# Importar correos desde PST a buzón de Outlook/Exchange usando COM
param (
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

    [Parameter(Mandatory=$false)]
    [int]$ItemsPerMinute = 120,

    [Parameter(Mandatory=$false)]
    [int]$BurstSize = 20,

    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 5,

    [Parameter(Mandatory=$false)]
    [int]$InitialBackoffMs = 1000,

    [Parameter(Mandatory=$false)]
    [int]$MaxBackoffMs = 30000,

    [Parameter(Mandatory=$false)]
    [switch]$AdaptiveThrottling,

    [Parameter(Mandatory=$false)]
    [string[]]$IncludeFolders,

    [Parameter(Mandatory=$false)]
    [switch]$ListFolders,

    [Parameter(Mandatory=$false)]
    [switch]$Json,

    [Parameter(Mandatory=$false)]
    [switch]$Headless
)

$ErrorActionPreference = "SilentlyContinue"

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
            [switch]$Completed,
            [int]$Copied,
            [int]$Moved,
            [int]$Skipped,
            [int]$Failed
        )
        $payload = @{
            type = "progress"
            activity = $Activity
            status = $Status
            percent = $PercentComplete
            completed = [bool]$Completed
            copied = $Copied
            moved = $Moved
            skipped = $Skipped
            failed = $Failed
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

function Get-SafeSubject {
    param($item)
    try { return [string]$item.Subject } catch { return "(unknown)" }
}

function Normalize-FolderPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return $Path.Trim().Replace("/", "\").Trim("\").ToLowerInvariant()
}

function Should-ProcessFolder {
    param([string]$FolderPath, [string[]]$SelectedFolders)
    if (-not $SelectedFolders -or $SelectedFolders.Count -eq 0) { return $true }
    $fp = Normalize-FolderPath $FolderPath
    foreach ($sel in $SelectedFolders) {
        if ($fp -eq $sel -or $fp.StartsWith("$sel\")) { return $true }
    }
    return $false
}

function Has-SelectedDescendant {
    param([string]$FolderPath, [string[]]$SelectedFolders)
    if (-not $SelectedFolders -or $SelectedFolders.Count -eq 0) { return $false }
    $fp = Normalize-FolderPath $FolderPath
    foreach ($sel in $SelectedFolders) {
        if ($sel.StartsWith("$fp\")) { return $true }
    }
    return $false
}

function Collect-PstFoldersRecursive {
    param($folder, [string]$pathPrefix, [ref]$out)
    try {
        $folderPath = if ($pathPrefix) { "$pathPrefix\$($folder.Name)" } else { $folder.Name }
        $count = 0
        try { $count = [int]$folder.Items.Count } catch {}
        $out.Value += [pscustomobject]@{
            type = "folder"
            path = $folderPath
            itemCount = $count
        }
        foreach ($sub in (Get-SubFolders-Safe -parentFolder $folder)) {
            Collect-PstFoldersRecursive -folder $sub -pathPrefix $folderPath -out $out
        }
    } catch {}
}

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

if ($FilterOnlyYear -and ($FilterOnlyYear -lt 1900 -or $FilterOnlyYear -gt 9999)) {
    Emit-ErrorPayload "FilterOnlyYear inválido: $FilterOnlyYear. Debe estar entre 1900 y 9999."
    exit 1
}

function Resolve-MonthTokenToNumber {
    param([string]$Token)

    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }

    $t = $Token.Trim().ToLower()
    $t = $t.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
        $unicode = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($unicode -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $t = $sb.ToString()

    switch ($t) {
        "1" { return 1 }
        "2" { return 2 }
        "3" { return 3 }
        "4" { return 4 }
        "5" { return 5 }
        "6" { return 6 }
        "7" { return 7 }
        "8" { return 8 }
        "9" { return 9 }
        "10" { return 10 }
        "11" { return 11 }
        "12" { return 12 }
        "enero" { return 1 }
        "ene" { return 1 }
        "january" { return 1 }
        "jan" { return 1 }
        "febrero" { return 2 }
        "feb" { return 2 }
        "february" { return 2 }
        "marzo" { return 3 }
        "mar" { return 3 }
        "march" { return 3 }
        "abril" { return 4 }
        "abr" { return 4 }
        "april" { return 4 }
        "mayo" { return 5 }
        "may" { return 5 }
        "junio" { return 6 }
        "jun" { return 6 }
        "june" { return 6 }
        "julio" { return 7 }
        "jul" { return 7 }
        "july" { return 7 }
        "agosto" { return 8 }
        "ago" { return 8 }
        "august" { return 8 }
        "aug" { return 8 }
        "septiembre" { return 9 }
        "setiembre" { return 9 }
        "sep" { return 9 }
        "set" { return 9 }
        "september" { return 9 }
        "octubre" { return 10 }
        "oct" { return 10 }
        "october" { return 10 }
        "noviembre" { return 11 }
        "nov" { return 11 }
        "november" { return 11 }
        "diciembre" { return 12 }
        "dic" { return 12 }
        "december" { return 12 }
        "dec" { return 12 }
        default { return $null }
    }
}

$FilterOnlyMonthNumbers = @()
if (-not [string]::IsNullOrWhiteSpace($FilterOnlyMonths)) {
    $tokens = $FilterOnlyMonths -split "," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (@($tokens).Count -eq 0) {
        Emit-ErrorPayload "FilterOnlyMonths no contiene meses válidos."
        exit 1
    }

    $resolved = @()
    foreach ($tk in $tokens) {
        $m = Resolve-MonthTokenToNumber -Token $tk
        if (-not $m) {
            Emit-ErrorPayload "Mes inválido en FilterOnlyMonths: '$tk'. Use nombres (enero..diciembre) o números (1..12)."
            exit 1
        }
        $resolved += [int]$m
    }

    $FilterOnlyMonthNumbers = @($resolved | Sort-Object -Unique)
}

$SelectedFolderFilters = @()
if ($IncludeFolders -and @($IncludeFolders).Count -gt 0) {
    foreach ($f in $IncludeFolders) {
        $nf = Normalize-FolderPath $f
        if (-not [string]::IsNullOrWhiteSpace($nf)) {
            $SelectedFolderFilters += $nf
        }
    }
    $SelectedFolderFilters = @($SelectedFolderFilters | Sort-Object -Unique)
}

$script:TokenBucket = @{
    capacity    = [double]$BurstSize
    tokens      = [double]$BurstSize
    refillRate  = [double]($ItemsPerMinute / 60.0)
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

function Resolve-TargetTopFolder {
    param(
        $targetStore,
        $targetRoot,
        [string]$SourceTopName
    )

    $n = ($SourceTopName.Trim().ToLowerInvariant())

    # OlDefaultFolders (valores COM)
    # 3=DeletedItems, 4=Outbox, 5=SentMail, 6=Inbox, 16=Drafts, 23=Junk
    if ($n -in @("bandeja de entrada", "inbox")) {
        try { return $targetStore.GetDefaultFolder(6) } catch {}
    } elseif ($n -in @("elementos eliminados", "deleted items")) {
        try { return $targetStore.GetDefaultFolder(3) } catch {}
    } elseif ($n -in @("elementos enviados", "sent items")) {
        try { return $targetStore.GetDefaultFolder(5) } catch {}
    } elseif ($n -in @("borradores", "drafts")) {
        try { return $targetStore.GetDefaultFolder(16) } catch {}
    } elseif ($n -in @("correo no deseado", "junk email")) {
        try { return $targetStore.GetDefaultFolder(23) } catch {}
    } elseif ($n -in @("bandeja de salida", "outbox")) {
        try { return $targetStore.GetDefaultFolder(4) } catch {}
    }

    return Ensure-ChildFolder -parent $targetRoot -Name $SourceTopName
}

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

function Build-DuplicateIndex {
    param($targetFolder)
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    try {
        $table = $targetFolder.GetTable("")
        $table.Columns.RemoveAll()
        try { $table.Columns.Add("urn:schemas:mailheader:message-id") | Out-Null } catch {}
        while (-not $table.EndOfTable) {
            $row = $table.GetNextRow()
            try {
                $v = $row["urn:schemas:mailheader:message-id"]
                if ($v) { [void]$set.Add([string]$v) }
            } catch {}
        }
    } catch {}
    return $set
}

function Restore-FolderRecursive {
    param(
        $sourceFolder,
        $destFolder,
        [string]$pathPrefix,
        [ref]$stats
    )
    $folderPath = if ($pathPrefix) { "$pathPrefix\$($sourceFolder.Name)" } else { $sourceFolder.Name }
    $activity = if ($Action -ieq "Move") { "Moviendo" } else { "Copiando" }

    $processCurrent = Should-ProcessFolder -FolderPath $folderPath -SelectedFolders $SelectedFolderFilters
    $hasSelectedBelow = Has-SelectedDescendant -FolderPath $folderPath -SelectedFolders $SelectedFolderFilters
    if (-not $processCurrent -and -not $hasSelectedBelow) {
        return
    }

    $dupIndex = $null
    if ($SkipDuplicates) {
        Emit-Log "info" "Indexando duplicados en: $folderPath"
        $dupIndex = Build-DuplicateIndex -targetFolder $destFolder
    }

    $items = $sourceFolder.Items
    $itemCount = 0
    try { $itemCount = [int]$items.Count } catch {}

    Emit-Log "info" "$activity $itemCount ítems de: $folderPath"

    if ($processCurrent -and $itemCount -gt 0) {
        for ($i = $itemCount; $i -ge 1; $i--) {
            $item = $null
            try { $item = $items.Item($i) } catch { continue }
            if (-not $item) { continue }

            if ($FilterOnlyYear) {
                $itemDate = $null
                try { $itemDate = $item.ReceivedTime } catch {}
                if (-not $itemDate) {
                    try { $itemDate = $item.CreationTime } catch {}
                }
                if (-not $itemDate -or $itemDate.Year -ne $FilterOnlyYear) {
                    $stats.Value.skipped++
                    continue
                }
            }

            if ($FilterOnlyMonthNumbers.Count -gt 0) {
                if (-not $itemDate) {
                    try { $itemDate = $item.ReceivedTime } catch {}
                    if (-not $itemDate) {
                        try { $itemDate = $item.CreationTime } catch {}
                    }
                }
                if (-not $itemDate -or ($itemDate.Month -notin $FilterOnlyMonthNumbers)) {
                    $stats.Value.skipped++
                    continue
                }
            }

            if ($dupIndex -and $dupIndex.Count -gt 0) {
                $msgId = $null
                try { $msgId = $item.PropertyAccessor.GetProperty("urn:schemas:mailheader:message-id") } catch {}
                if ($msgId -and $dupIndex.Contains([string]$msgId)) {
                    $stats.Value.skipped++
                    continue
                }
            }

            $itemSize = 0
            try { $itemSize = [long]$item.Size } catch {}
            if ($itemSize -gt 157286400) {
                $stats.Value.failed++
                $stats.Value.failures += @{
                    folder = $folderPath
                    subject = (Get-SafeSubject $item)
                    reason = "too_large"
                    sizeBytes = $itemSize
                }
                Emit-Log "warn" "Ítem > 150MB ignorado: $(Get-SafeSubject $item)"
                continue
            }

            Wait-ForToken

            try {
                Invoke-WithRetry -OperationName "$Action ítem" -Operation {
                    if ($Action -ieq "Move") {
                        [void]$item.Move($destFolder)
                    } else {
                        $copied = $item.Copy()
                        if ($copied) {
                            [void]$copied.Move($destFolder)
                        } else {
                            throw "No se pudo copiar el ítem."
                        }
                    }
                }
                if ($Action -ieq "Move") { $stats.Value.moved++ } else { $stats.Value.copied++ }
            } catch {
                $stats.Value.failed++
                $reason = if (Is-ThrottlingError $_) { "throttled_max_retries" } else { "error" }
                $stats.Value.failures += @{
                    folder = $folderPath
                    subject = (Get-SafeSubject $item)
                    reason = $reason
                    message = "$($_.Exception.Message)"
                }
                Emit-Log "error" "Falló ítem en $folderPath : $($_.Exception.Message)"
            }

            $stats.Value.processed++
            $pct = 0
            if ($stats.Value.total -gt 0) {
                $pct = [int][Math]::Round(($stats.Value.processed / $stats.Value.total) * 100)
                if ($pct -gt 99 -and $stats.Value.processed -lt $stats.Value.total) { $pct = 99 }
                if ($pct -eq 0 -and $stats.Value.processed -gt 0) { $pct = 1 }
            }
            Write-Progress -Activity "Restauracion PST -> Buzon" -Status "$folderPath ($($stats.Value.processed)/$($stats.Value.total))" -PercentComplete $pct -Copied $stats.Value.copied -Moved $stats.Value.moved -Skipped $stats.Value.skipped -Failed $stats.Value.failed

            Emit-ThrottleStats
        }
    }

    foreach ($sub in (Get-SubFolders-Safe -parentFolder $sourceFolder)) {
        $destSub = Ensure-ChildFolder -parent $destFolder -Name $sub.Name
        Restore-FolderRecursive -sourceFolder $sub -destFolder $destSub -pathPrefix $folderPath -stats $stats
    }
}

if (-not $PstPath) { Emit-ErrorPayload "Se requiere -PstPath."; exit 1 }
if (-not (Test-Path $PstPath)) { Emit-ErrorPayload "PST no encontrado: $PstPath"; exit 1 }

$namespace = Get-OutlookNamespace

$alreadyMounted = $false
$pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
if ($pstStore) {
    $alreadyMounted = $true
    Emit-Log "info" "PST ya estaba montado."
} else {
    Emit-Log "info" "Montando PST: $PstPath"
    try { $namespace.AddStoreEx($PstPath, 3) } catch {
        Emit-ErrorPayload "No se pudo montar el PST: $($_.Exception.Message)"
        exit 1
    }
    $pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
    if (-not $pstStore) {
        Emit-ErrorPayload "PST montado pero no localizado."
        exit 1
    }
}

$pstRoot = $pstStore.GetRootFolder()

if ($ListFolders) {
    $flat = [ref]@()
    foreach ($tf in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
        Collect-PstFoldersRecursive -folder $tf -pathPrefix "" -out $flat
    }

    Write-Output (@{ type = "folders"; count = @($flat.Value).Count } | ConvertTo-Json -Compress -Depth 6)
    foreach ($f in $flat.Value) {
        Write-Output ($f | ConvertTo-Json -Compress -Depth 6)
    }

    if (-not $alreadyMounted) {
        try { $namespace.RemoveStore($pstRoot) } catch {}
    }
    exit 0
}

if (-not $TargetStoreId) { Emit-ErrorPayload "Se requiere -TargetStoreId."; exit 1 }

$targetStore = Get-StoreByIdOrPath -namespace $namespace -StoreId $TargetStoreId
if (-not $targetStore) {
    Emit-ErrorPayload "No se encontró el buzón destino (StoreId=$TargetStoreId)."
    exit 1
}
Emit-Log "info" "Buzón destino: $($targetStore.DisplayName)"

Emit-Log "info" "Contando ítems del PST (para progreso)..."
$totalItems = [ref]0
$totalBytes = [ref]([long]0)
$folderList = [ref]@()
$pstRoot = $pstStore.GetRootFolder()
foreach ($tf in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
    Analyze-PstFolderRecursive -folder $tf -pathPrefix "" -totalItems $totalItems -totalBytes $totalBytes -folderList $folderList
}
Emit-Log "info" "Total ítems a procesar: $($totalItems.Value)"

$stats = [ref]@{
    copied = 0
    moved = 0
    skipped = 0
    failed = 0
    processed = 0
    total = [int]$totalItems.Value
    failures = @()
}

$startTime = [DateTime]::UtcNow
$targetRoot = $targetStore.GetRootFolder()

foreach ($sourceTop in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
    $destTop = Resolve-TargetTopFolder -targetStore $targetStore -targetRoot $targetRoot -SourceTopName $sourceTop.Name
    Restore-FolderRecursive -sourceFolder $sourceTop -destFolder $destTop -pathPrefix "" -stats $stats
}

Write-Progress -Activity "Restauracion PST -> Buzon" -Status "Completado" -PercentComplete 100 -Completed -Copied $stats.Value.copied -Moved $stats.Value.moved -Skipped $stats.Value.skipped -Failed $stats.Value.failed
Emit-ThrottleStats -Force

if (-not $alreadyMounted) {
    try { $namespace.RemoveStore($pstRoot) } catch {
        Emit-Log "warn" "No se pudo desmontar el PST: $($_.Exception.Message)"
    }
}

$elapsed = [long](([DateTime]::UtcNow - $startTime).TotalMilliseconds)
$payload = @{
    type = "restoreResult"
    filterOnlyYear = if ($FilterOnlyYear) { [int]$FilterOnlyYear } else { $null }
    filterOnlyMonths = if ($FilterOnlyMonthNumbers.Count -gt 0) { @($FilterOnlyMonthNumbers) } else { @() }
    copied = [int]$stats.Value.copied
    moved = [int]$stats.Value.moved
    skipped = [int]$stats.Value.skipped
    failed = [int]$stats.Value.failed
    elapsedMs = $elapsed
    throttleEvents = [int]$script:TokenBucket.throttleErrors
    totalWaitedMs = [long]$script:TokenBucket.totalWaitedMs
    failures = @($stats.Value.failures)
}
Write-Output ($payload | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue)
exit 0
