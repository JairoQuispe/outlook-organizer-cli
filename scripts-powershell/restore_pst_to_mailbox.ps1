# Restaurar correos desde un archivo PST hacia un buzón de Exchange Online
# Vía Outlook COM (MAPI RPC/HTTP). Respeta throttling EXO con token bucket + backoff.
param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("list-stores", "analyze-pst", "run-restore")]
    [string]$Command = "list-stores",

    [Parameter(Mandatory=$false)]
    [string]$PstPath,

    [Parameter(Mandatory=$false)]
    [string]$TargetStoreId,

    [Parameter(Mandatory=$false)]
    [ValidateSet("Copy", "Move", "copy", "move")]
    [string]$Action = "Copy",

    [Parameter(Mandatory=$false)]
    [int]$FilterOnlyYear,

    [Parameter(Mandatory=$false)]
    [string]$FilterOnlyMonths,

    [Parameter(Mandatory=$false)]
    [switch]$SkipDuplicates,

    # --- Throttling (defaults conservadores) ---
    [Parameter(Mandatory=$false)]
    [int]$ItemsPerMinute = 30,

    [Parameter(Mandatory=$false)]
    [int]$BurstSize = 10,

    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 5,

    [Parameter(Mandatory=$false)]
    [int]$InitialBackoffMs = 1000,

    [Parameter(Mandatory=$false)]
    [int]$MaxBackoffMs = 30000,

    [Parameter(Mandatory=$false)]
    [switch]$AdaptiveThrottling,

    [Parameter(Mandatory=$false)]
    [switch]$Json,

    [Parameter(Mandatory=$false)]
    [switch]$Headless,

    [Parameter(Mandatory=$false)]
    [string]$LogDir
)

# --- Overrides JSON (shape compatible con parser Rust) ---
if ($Json -or $Headless) {
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

    function Write-Host {
        param(
            [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
            [object[]]$Object,
            [ConsoleColor]$ForegroundColor,
            [ConsoleColor]$BackgroundColor,
            [switch]$NoNewline,
            [object]$Separator
        )
        $msg = ($Object | ForEach-Object { "$_" }) -join " "
        $payload = @{ type = "log"; level = "info"; message = $msg }
        if ($Json) { Write-Output ($payload | ConvertTo-Json -Compress -Depth 6) }
    }

    function Write-Warning {
        param([Parameter(ValueFromRemainingArguments = $true)] [object[]]$Message)
        $msg = ($Message | ForEach-Object { "$_" }) -join " "
        $payload = @{ type = "log"; level = "warn"; message = $msg }
        if ($Json) { Write-Output ($payload | ConvertTo-Json -Compress -Depth 6) }
    }

    function Write-Progress {
        param(
            [string]$Activity,
            [string]$Status,
            [int]$PercentComplete,
            [switch]$Completed
        )
        $payload = @{
            type = "progress"
            activity = $Activity
            status = $Status
            percent = $PercentComplete
            completed = [bool]$Completed
        }
        if ($Json) { Write-Output ($payload | ConvertTo-Json -Compress -Depth 6) }
    }
}

function Emit-Log {
    param([string]$Level, [string]$Message)
    $payload = @{ type = "log"; level = $Level; message = $Message }
    if ($Json) {
        Write-Output ($payload | ConvertTo-Json -Compress -Depth 6)
    } else {
        Write-Host "[$Level] $Message"
    }
}

function Emit-ErrorPayload {
    param([string]$Message)
    $payload = @{ type = "error"; message = $Message }
    if ($Json) {
        Write-Output ($payload | ConvertTo-Json -Compress -Depth 6)
    } else {
        Write-Error $Message
    }
}

function Format-Bytes {
    param ($bytes)
    if ($bytes -ge 1GB) { return "{0:N2} GB" -f ($bytes / 1GB) }
    elseif ($bytes -ge 1MB) { return "{0:N2} MB" -f ($bytes / 1MB) }
    elseif ($bytes -ge 1KB) { return "{0:N2} KB" -f ($bytes / 1KB) }
    else { return "$bytes Bytes" }
}

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$storeHelpersPath = Join-Path $scriptDir "outlook-list-stores.ps1"
if (Test-Path $storeHelpersPath) {
    . $storeHelpersPath
} else {
    Emit-ErrorPayload "No se encontró el archivo requerido: $storeHelpersPath"
    exit 1
}

# --- Token Bucket (rate limiter) ---
$script:TokenBucket = @{
    capacity    = [double]$BurstSize
    tokens      = [double]$BurstSize
    refillRate  = [double]($ItemsPerMinute / 60.0)  # tokens/segundo
    lastRefill  = [DateTime]::UtcNow
    totalWaitedMs = [long]0
    throttleErrors = [int]0
    adaptiveMultiplier = [double]1.0
    lastStatsEmit = [DateTime]::UtcNow
}

function Wait-ForToken {
    while ($true) {
        $now = [DateTime]::UtcNow
        $elapsed = ($now - $script:TokenBucket.lastRefill).TotalSeconds
        $effectiveRate = $script:TokenBucket.refillRate * $script:TokenBucket.adaptiveMultiplier
        $script:TokenBucket.tokens = [Math]::Min(
            $script:TokenBucket.capacity,
            $script:TokenBucket.tokens + ($elapsed * $effectiveRate)
        )
        $script:TokenBucket.lastRefill = $now

        if ($script:TokenBucket.tokens -ge 1.0) {
            $script:TokenBucket.tokens -= 1.0
            return
        }

        $needed = 1.0 - $script:TokenBucket.tokens
        $waitMs = [int](($needed / [Math]::Max($effectiveRate, 0.001)) * 1000)
        $waitMs = [Math]::Max(50, [Math]::Min($waitMs, 5000))
        $script:TokenBucket.totalWaitedMs += $waitMs
        Start-Sleep -Milliseconds $waitMs
    }
}

function Emit-ThrottleStats {
    param([switch]$Force)
    $now = [DateTime]::UtcNow
    if (-not $Force) {
        if (($now - $script:TokenBucket.lastStatsEmit).TotalSeconds -lt 10) { return }
    }
    $script:TokenBucket.lastStatsEmit = $now
    $effRate = $script:TokenBucket.refillRate * $script:TokenBucket.adaptiveMultiplier * 60.0
    $payload = @{
        type = "throttleStats"
        effectiveRate = [double]([Math]::Round($effRate, 2))
        currentTokens = [double]([Math]::Round($script:TokenBucket.tokens, 2))
        burstCapacity = [int]$script:TokenBucket.capacity
        totalWaitedMs = [long]$script:TokenBucket.totalWaitedMs
        throttleErrors = [int]$script:TokenBucket.throttleErrors
        adaptiveMultiplier = [double]([Math]::Round($script:TokenBucket.adaptiveMultiplier, 3))
    }
    if ($Json) { Write-Output ($payload | ConvertTo-Json -Compress -Depth 6) }
}

function Is-ThrottlingError {
    param($errorRecord)
    $msg = "$errorRecord"
    if ($msg -match "0x80040115" -or $msg -match "0x8004011D") { return $true }
    if ($msg -match "Server Busy" -or $msg -match "throttl" -or $msg -match "budget" -or $msg -match "too many requests" -or $msg -match "429") { return $true }
    return $false
}

function Invoke-WithRetry {
    param(
        [ScriptBlock]$Operation,
        [string]$OperationName = "operation"
    )
    $attempt = 0
    while ($true) {
        try {
            return & $Operation
        } catch {
            if (-not (Is-ThrottlingError $_) -or $attempt -ge $MaxRetries) {
                throw
            }
            $script:TokenBucket.throttleErrors++
            $delay = [Math]::Min(
                $InitialBackoffMs * [Math]::Pow(2, $attempt),
                $MaxBackoffMs
            )
            $delay = [int]$delay
            Emit-Log "warn" "Throttling en $OperationName, backoff ${delay}ms (intento $($attempt+1)/$MaxRetries): $($_.Exception.Message)"
            if ($AdaptiveThrottling) {
                $script:TokenBucket.adaptiveMultiplier = [Math]::Max(0.1, $script:TokenBucket.adaptiveMultiplier * 0.7)
                $newRate = [int]($script:TokenBucket.refillRate * $script:TokenBucket.adaptiveMultiplier * 60.0)
                Emit-Log "info" "Adaptive throttling: nueva tasa efectiva ~${newRate} items/min"
            }
            Start-Sleep -Milliseconds $delay
            $attempt++
        }
    }
}

# --- Outlook COM ---
function Get-OutlookNamespace {
    try {
        $outlook = $null
        $namespace = $null

        try {
            $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $namespace = $outlook.GetNamespace("MAPI")
            try { $namespace.Logon("", "", $false, $false) } catch {}
        } catch {
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            try { $namespace.Logon("", "", $false, $true) } catch {}
        }

        return $namespace
    } catch {
        Emit-ErrorPayload "No se pudo iniciar Outlook: $($_.Exception.Message)"
        throw
    }
}

function Get-StoreByIdOrPath {
    param($namespace, [string]$StoreId, [string]$FilePath)
    foreach ($s in $namespace.Stores) {
        if ($StoreId -and ($s.StoreID -eq $StoreId)) { return $s }
        if ($FilePath) {
            try { if ($s.FilePath -and ($s.FilePath -eq $FilePath)) { return $s } } catch {}
        }
    }
    return $null
}

function Get-SubFolders-Safe {
    param($parentFolder)
    try { return $parentFolder.Folders } catch { return @() }
}

function Ensure-ChildFolder {
    param($parent, [string]$Name)
    foreach ($f in (Get-SubFolders-Safe -parentFolder $parent)) {
        if ($f.Name -ieq $Name) { return $f }
    }
    return $parent.Folders.Add($Name)
}

# ==========================================================================
# COMMAND: list-stores
# ==========================================================================
if ($Command -eq "list-stores") {
    $namespace = Get-OutlookNamespace
    $stores = @(Get-ConnectedOutlookStores -namespace $namespace)

    if ($stores.Count -eq 0) {
        $storesJson = "[]"
    } elseif ($stores.Count -eq 1) {
        $singleJson = $stores[0] | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue
        $storesJson = "[$singleJson]"
    } else {
        $storesJson = $stores | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue
    }

    Write-Output ('{"type":"stores","stores":' + $storesJson + '}')
    exit 0
}

# ==========================================================================
# COMMAND: analyze-pst
# ==========================================================================
function Analyze-PstFolderRecursive {
    param($folder, [string]$pathPrefix, [ref]$totalItems, [ref]$totalBytes, [ref]$folderList)
    try {
        $folderPath = if ($pathPrefix) { "$pathPrefix\$($folder.Name)" } else { $folder.Name }
        $itemCount = 0
        $sizeBytes = [long]0

        try {
            $items = $folder.Items
            $itemCount = [int]$items.Count
            try {
                $table = $folder.GetTable("")
                $table.Columns.RemoveAll()
                $table.Columns.Add("Size") | Out-Null
                while (-not $table.EndOfTable) {
                    $row = $table.GetNextRow()
                    try { $sizeBytes += [long]$row["Size"] } catch {}
                }
            } catch {}
        } catch {}

        $totalItems.Value += $itemCount
        $totalBytes.Value += $sizeBytes
        $folderList.Value += [pscustomobject]@{
            path = $folderPath
            itemCount = $itemCount
            sizeBytes = $sizeBytes
            sizeHuman = (Format-Bytes -bytes $sizeBytes)
        }

        foreach ($sub in (Get-SubFolders-Safe -parentFolder $folder)) {
            Analyze-PstFolderRecursive -folder $sub -pathPrefix $folderPath -totalItems $totalItems -totalBytes $totalBytes -folderList $folderList
        }
    } catch {
        Emit-Log "warn" "Error analizando carpeta '$($folder.Name)': $($_.Exception.Message)"
    }
}

if ($Command -eq "analyze-pst") {
    if (-not $PstPath) {
        Emit-ErrorPayload "Se requiere -PstPath para analyze-pst."
        exit 1
    }
    if (-not (Test-Path $PstPath)) {
        Emit-ErrorPayload "PST no encontrado: $PstPath"
        exit 1
    }

    $namespace = Get-OutlookNamespace
    Emit-Log "info" "Montando PST: $PstPath"

    $alreadyMounted = $false
    $pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
    if ($pstStore) {
        $alreadyMounted = $true
        Emit-Log "info" "PST ya estaba montado en Outlook."
    } else {
        try {
            $namespace.AddStoreEx($PstPath, 3)  # 3 = olStoreUnicode
        } catch {
            Emit-ErrorPayload "No se pudo montar el PST: $($_.Exception.Message)"
            exit 1
        }
        $pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
        if (-not $pstStore) {
            Emit-ErrorPayload "PST montado pero no localizado en \$namespace.Stores."
            exit 1
        }
    }

    $totalItems = [ref]0
    $totalBytes = [ref]([long]0)
    $folderList = [ref]@()

    $root = $pstStore.GetRootFolder()
    $topFolders = Get-SubFolders-Safe -parentFolder $root
    $topCount = @($topFolders).Count
    $idx = 0
    foreach ($tf in $topFolders) {
        $idx++
        $percent = if ($topCount -gt 0) { [int][Math]::Round(($idx / $topCount) * 100) } else { 100 }
        if ($percent -gt 99 -and $idx -lt $topCount) { $percent = 99 }
        Write-Progress -Activity "Análisis PST" -Status "Carpeta $idx/$topCount : $($tf.Name)" -PercentComplete $percent
        Analyze-PstFolderRecursive -folder $tf -pathPrefix "" -totalItems $totalItems -totalBytes $totalBytes -folderList $folderList
    }
    Write-Progress -Activity "Análisis PST" -Status "Completado" -PercentComplete 100 -Completed

    if (-not $alreadyMounted) {
        try { $namespace.RemoveStore($root) } catch {
            Emit-Log "warn" "No se pudo desmontar el PST tras análisis: $($_.Exception.Message)"
        }
    }

    $payload = @{
        type = "pstAnalysis"
        pstPath = $PstPath
        totalItems = [int]$totalItems.Value
        totalSizeBytes = [long]$totalBytes.Value
        totalSizeHuman = (Format-Bytes -bytes $totalBytes.Value)
        folders = @($folderList.Value)
    }
    Write-Output ($payload | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue)
    exit 0
}

# ==========================================================================
# COMMAND: run-restore
# ==========================================================================

if ($Command -eq "run-restore") {
    $importScriptPath = Join-Path $scriptDir "outlook-import-pst.ps1"
    if (-not (Test-Path $importScriptPath)) {
        Emit-ErrorPayload "No se encontró el archivo requerido: $importScriptPath"
        exit 1
    }

    $args = @(
        "-ExecutionPolicy", "Bypass",
        "-File", $importScriptPath,
        "-PstPath", $PstPath,
        "-TargetStoreId", $TargetStoreId,
        "-Action", $Action,
        "-FilterOnlyYear", $FilterOnlyYear,
        "-FilterOnlyMonths", $FilterOnlyMonths,
        "-ItemsPerMinute", $ItemsPerMinute,
        "-BurstSize", $BurstSize,
        "-MaxRetries", $MaxRetries,
        "-InitialBackoffMs", $InitialBackoffMs,
        "-MaxBackoffMs", $MaxBackoffMs
    )

    if ($SkipDuplicates) { $args += "-SkipDuplicates" }
    if ($AdaptiveThrottling) { $args += "-AdaptiveThrottling" }
    if ($Json) { $args += "-Json" }
    if ($Headless) { $args += "-Headless" }

    & powershell @args
    exit $LASTEXITCODE
}

Emit-ErrorPayload "Comando no soportado: $Command"
exit 1
