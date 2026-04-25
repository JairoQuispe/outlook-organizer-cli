# Define backup year and target path as parameters
param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("interactive", "list-stores", "analyze-store", "run-backup")]
    [string]$Command = "interactive",

    [Parameter(Mandatory=$false)]
    [string]$StoreId,

    [Parameter(Mandatory=$false)]
    [string]$BackupYear,

    [Parameter(Mandatory=$false)]
    [string]$BackupMonths,
    
    [Parameter(Mandatory=$false)]
    [string]$TargetBackupPath = "C:\Backups",

    [Parameter(Mandatory=$false)]
    [ValidateSet("Copy", "Move", "copy", "move")]
    [string]$BackupAction = "Copy",

    [Parameter(Mandatory=$false)]
    [string]$PstName,

    [Parameter(Mandatory=$false)]
    [switch]$Json,

    [Parameter(Mandatory=$false)]
    [switch]$Headless,

    [Parameter(Mandatory=$false)]
    [string]$LogDir
)

# --- Dynamic Year Resolution ---
if ($BackupYear -eq "PREVIOUS") {
    $BackupYear = (Get-Date).AddYears(-1).Year
} elseif ($BackupYear -eq "CURRENT") {
    $BackupYear = (Get-Date).Year
}

$backupYearProvided = -not [string]::IsNullOrWhiteSpace($BackupYear)

# Cast to int for internal use (only when provided)
if ($backupYearProvided) {
    try {
        $BackupYearInt = [int]$BackupYear
        $BackupYear = $BackupYearInt
    } catch {
        if ($Command -eq "run-backup") {
            Write-Error "BackupYear inválido: $BackupYear"
            exit 1
        }
    }
} elseif ($Command -eq "run-backup") {
    Write-Error "BackupYear es requerido para run-backup."
    exit 1
}

# Validate Year Range for run-backup
if ($Command -eq "run-backup" -and ($BackupYear -lt 1900 -or $BackupYear -gt 9999)) {
    Write-Error "El año debe estar entre 1900 y 9999. Valor recibido: $BackupYear"
    exit 1
}

# --- Headless Logging Setup ---
$LogFile = $null
if ($Headless) {
    if (-not $LogDir) {
        $LogDir = Join-Path $env:LOCALAPPDATA "OutlookBackup\Logs"
    }
    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
    }
    # Log filename: history-YYYY-MM-DD_HH-mm-ss.json
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $LogFile = Join-Path $LogDir "history-$timestamp.json"
    
    # Initialize Log File with metadata
    $meta = @{
        taskName = "OutlookBackup" # Placeholder, improved later
        timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        status = "Running"
        details = @{
            year = $BackupYear
            action = $BackupAction
        }
    }
    $meta | ConvertTo-Json -Compress | Out-File -FilePath $LogFile -Encoding UTF8
}

# Global Trap for Headless Errors
trap {
    if ($Headless -and $LogFile) {
        $errMeta = @{
            taskName = "OutlookBackup"
            timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            status = "Error"
            message = "Script Error: $($_.Exception.Message)"
            details = @{ errorItem = "$_" }
        }
        $errMeta | ConvertTo-Json -Compress | Out-File -FilePath $LogFile -Append -Encoding UTF8 -Force
    }
    # Allow normal error propagation
}

# Helper to append to log file safely
function Append-Log {
    param($payload)
    if ($Headless -and $LogFile) {
        # We append line-delimited JSON or just overwrite status?
        # For history, we want the FINAL status mostly, but streaming logs helps debugging.
        # Let's append log lines.
        # However, a valid JSON file is one object.
        # We can write a "log" file that is line-delimited JSON (NDJSON).
        $json = $payload | ConvertTo-Json -Compress -Depth 6
        $json | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
}

# Set the execution policy for this session to bypass for this script
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

# --- Parse Parameters ---

# Lista de carpetas a excluir del analisis y backup (Nombre exacto o patrones comunes)
$excludedFolderNames = @("Contactos", "Contacts", "Calendar", "Calendario", "Journal", "Diario", "Tasks", "Tareas", "Notes", "Notas", "Sync Issues", "Problemas de sincronización", "Yammer Root", "Quick Step Settings", "Conversation Action Settings", "Recipient Cache")

# Convert comma-separated string to int array if needed
$BackupMonthsList = @()
if ($BackupMonths) {
    try {
        if ($BackupMonths -is [string]) {
            $BackupMonthsList = $BackupMonths -split "," | ForEach-Object { [int]$_.Trim() }
        } else {
            $BackupMonthsList = $BackupMonths
        }
    } catch {
        Write-Error "Error parsing BackupMonths: $_"
    }
}

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
        
        if ($Json) {
            Write-Output ($payload | ConvertTo-Json -Compress -Depth 6)
        }
        if ($Headless) {
             # Don't clutter log file with every info message, maybe just key ones?
             # For now, log everything to help debug headless.
             # Append-Log $payload
        }
    }

    function Write-Warning {
        param([Parameter(ValueFromRemainingArguments = $true)] [object[]]$Message)
        $msg = ($Message | ForEach-Object { "$_" }) -join " "
        $payload = @{ type = "log"; level = "warn"; message = $msg }
        if ($Json) {
            Write-Output ($payload | ConvertTo-Json -Compress -Depth 6)
        }
        Append-Log $payload
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
        if ($Json) {
            Write-Output ($payload | ConvertTo-Json -Compress -Depth 6)
        }
        # Progress is too verbose for log file
    }
}

function Format-Bytes {
    param ($bytes)
    if ($bytes -ge 1GB) {
        return "{0:N2} GB" -f ($bytes / 1GB)
    } elseif ($bytes -ge 1MB) {
        return "{0:N2} MB" -f ($bytes / 1MB)
    } elseif ($bytes -ge 1KB) {
        return "{0:N2} KB" -f ($bytes / 1KB)
    } else {
        return "$bytes Bytes"
    }
}

function Get-RomanNumeral {
    param ([int]$number)
    if ($number -lt 1 -or $number -gt 3999) { return $number.ToString() }
    $dict = @{ 1000="M"; 900="CM"; 500="D"; 400="CD"; 100="C"; 90="XC"; 50="L"; 40="XL"; 10="X"; 9="IX"; 5="V"; 4="IV"; 1="I" }
    $keys = $dict.Keys | Sort-Object -Descending
    $result = ""
    foreach ($key in $keys) {
        while ($number -ge $key) {
            $result += $dict[$key]
            $number -= $key
        }
    }
    return $result
}

function Get-SubFolders-Safe {
    param (
        $parentFolder
    )
    $maxRetries = 6
    $baseDelay = 4

    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            return $parentFolder.Folders
        } catch {
            $err = $_.Exception.Message
            $hresult = $_.Exception.HResult
            # Check for typical connection errors (RPC server unavailable, Network problems)
            # 0x80040115 (MAPI_E_NETWORK_ERROR), 0x800706BA (RPC_S_SERVER_UNAVAILABLE), etc.
            
            if ($i -lt $maxRetries) {
                Write-Host "  [WARN] Fallo al acceder a subcarpetas de '$($parentFolder.Name)'. Intento $i/$maxRetries. Reintentando en $($baseDelay * $i)s... (Error: $err)" -ForegroundColor Yellow
                Start-Sleep -Seconds ($baseDelay * $i)
            } else {
                Write-Host "  [ERROR] No se pudieron leer las subcarpetas de '$($parentFolder.Name)' tras múltiples intentos. Saltando rama. (Error: $err)" -ForegroundColor Red
            }
        }
    }
    return @()
}


function Get-YearDistributionRecursive {
    param (
        $folder,
        [hashtable]$yearCounts,
        [hashtable]$yearSizes,
        [hashtable]$monthCounts,
        [hashtable]$monthSizes,
        [hashtable]$folderStats,
        [int]$yearFilter = 0,
        [int[]]$monthsFilter
    )

    $cnt = 0

    # Exclusion Check
    if ($excludedFolderNames -contains $folder.Name) {
        # Write-Host "  [DEBUG] Excluyendo carpeta: $($folder.Name)" -ForegroundColor DarkGray
        return 0
    }

    try {
        if ($folder.Items.Count -gt 0) {
            try {
                $filter = ""
                # Si hay filtro de año, aplicamos filtro MAPI para acelerar
                if ($yearFilter -gt 1900) {
                    $startDate = Get-Date -Year $yearFilter -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                    $endDate = Get-Date -Year ($yearFilter + 1) -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                    # Filtro SQL-like para GetTable
                    # [ReceivedTime] >= '...' AND [ReceivedTime] < '...'
                    # Ojo: ReceivedTime puede ser null en borradores, usar CreationTime como fallback es mas complejo en filtro directo.
                    # Sin embargo, GetTable con filtro es mucho mas rapido.
                    # Si falla el filtro, cae al catch y hace iteración manual, lo cual está bien.
                    $filter = "[ReceivedTime] >= '$($startDate.ToString('MM/dd/yyyy HH:mm tt'))' AND [ReceivedTime] < '$($endDate.ToString('MM/dd/yyyy HH:mm tt'))'"
                }

                $table = $folder.GetTable($filter)
                $table.Columns.RemoveAll()
                $table.Columns.Add("ReceivedTime") | Out-Null
                $table.Columns.Add("CreationTime") | Out-Null
                $table.Columns.Add("Size") | Out-Null

                while (-not $table.EndOfTable) {
                    $rows = $table.GetNextRows(100)
                    foreach ($row in $rows) {
                        $date = $row["ReceivedTime"]
                        if (-not $date) { $date = $row["CreationTime"] }
                        if ($date) {
                            $y = $date.Year
                            $m = $date.Month
                            
                            # Filtros adicionales si GetTable no pudo filtrar todo (ej. CreationTime o Month)
                            if ($yearFilter -gt 1900 -and $y -ne $yearFilter) { continue }
                            if ($monthsFilter -and $m -notin $monthsFilter) { continue }

                            $s = $row["Size"]
                            
                            # Year stats
                            if ($yearCounts.ContainsKey($y)) {
                                $yearCounts[$y]++
                                $yearSizes[$y] += $s
                            } else {
                                $yearCounts[$y] = 1
                                $yearSizes[$y] = $s
                            }

                            # Folder stats
                            if ($folderStats) {
                                $fPath = $folder.FolderPath
                                if (-not $folderStats.ContainsKey($fPath)) { $folderStats[$fPath] = @{} }
                                if ($folderStats[$fPath].ContainsKey($y)) { $folderStats[$fPath][$y]++ } else { $folderStats[$fPath][$y] = 1 }
                            }

                            # Month stats (Key: "Year-Month")
                            $ymKey = "$y-$m"
                            if ($monthCounts.ContainsKey($ymKey)) {
                                $monthCounts[$ymKey]++
                                $monthSizes[$ymKey] += $s
                            } else {
                                $monthCounts[$ymKey] = 1
                                $monthSizes[$ymKey] = $s
                            }

                            $cnt++
                        }
                    }
                }
            } catch {
                # Fallback manual iteration
                foreach ($item in $folder.Items) {
                    $date = $item.ReceivedTime
                    if (-not $date) { $date = $item.CreationTime }
                    if ($date) {
                        $y = $date.Year
                        $m = $date.Month
                        
                        if ($yearFilter -gt 1900 -and $y -ne $yearFilter) { continue }
                        if ($monthsFilter -and $m -notin $monthsFilter) { continue }

                        $s = $item.Size
                        
                        # Year stats
                        if ($yearCounts.ContainsKey($y)) { $yearCounts[$y]++; $yearSizes[$y] += $s }
                        else { $yearCounts[$y] = 1; $yearSizes[$y] = $s }

                        # Folder stats
                        if ($folderStats) {
                            $fPath = $folder.FolderPath
                            if (-not $folderStats.ContainsKey($fPath)) { $folderStats[$fPath] = @{} }
                            if ($folderStats[$fPath].ContainsKey($y)) { $folderStats[$fPath][$y]++ } else { $folderStats[$fPath][$y] = 1 }
                        }

                        # Month stats
                        $ymKey = "$y-$m"
                        if ($monthCounts.ContainsKey($ymKey)) { $monthCounts[$ymKey]++; $monthSizes[$ymKey] += $s }
                        else { $monthCounts[$ymKey] = 1; $monthSizes[$ymKey] = $s }

                        $cnt++
                    }
                }
            }
        }
    } catch {
         Write-Host "  [WARN] Error al leer items de carpeta '$($folder.Name)': $($_.Exception.Message)" -ForegroundColor DarkGray
    }

    $subFolders = Get-SubFolders-Safe -parentFolder $folder
    if ($subFolders) {
        foreach ($sub in $subFolders) {
            $cnt += Get-YearDistributionRecursive -folder $sub -yearCounts $yearCounts -yearSizes $yearSizes -monthCounts $monthCounts -monthSizes $monthSizes -folderStats $folderStats -yearFilter $yearFilter -monthsFilter $monthsFilter
        }
    }

    return $cnt
}


# Create Outlook Application object
try {
    # Prefer the active Outlook instance to avoid session inconsistencies
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        Write-Host "Usando instancia activa de Outlook." -ForegroundColor Green
        
        # Aunque estemos conectados a una instancia activa, si está en modo "headless" o zombie (sin UI),
        # algunas llamadas pueden fallar si no hay una sesión MAPI explícita o si falta Logon.
        $namespace = $outlook.GetNamespace("MAPI")
        try {
            # Intentamos Logon para asegurar que la sesión esté 'viva' y conectada al perfil.
            # (ShowDialog=False, NewSession=False)
            $namespace.Logon("", "", $false, $false) 
        } catch {
            # Si falla el logon en una instancia activa, generalmente es ignorable porque ya debería tener sesión.
            # Pero si es una instancia "zombie", esto podría reactivarla.
        }
    } catch {
        Write-Host "Outlook no está activo o accesible. Intentando iniciar una nueva instancia..." -ForegroundColor Yellow
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        try {
            $namespace.Logon("", "", $false, $true)
        } catch {
            Write-Warning "No se pudo hacer Logon automático. Si Outlook pide perfil, la ventana podría estar oculta."
        }
    }
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    Write-Host "Error al conectar con Outlook: $($_.Exception.Message)"
    exit 1
}

Write-Host "Conectado a Outlook."

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$storeHelpersPath = Join-Path $scriptDir "outlook-list-stores.ps1"
if (Test-Path $storeHelpersPath) {
    . $storeHelpersPath
} else {
    throw "No se encontró el archivo requerido: $storeHelpersPath"
}

# Initialize sourceFolders
$sourceFolders = $null

if ($Command -eq "list-stores") {
    $result = Get-ConnectedOutlookStores -namespace $namespace

    if ($Json) {
        Write-Output (@{ type = "stores"; stores = $result } | ConvertTo-Json -Compress -Depth 6)
    } else {
        $i = 1
        foreach ($s in $result) {
            $pathDisplay = if ($s.filePath) { $s.filePath } else { "Modo Online / Sin ruta local" }
            Write-Host "[$i] $($s.displayName) - ($pathDisplay)"
            $i++
        }
    }
    exit 0
}

if ($Command -eq "analyze-store") {
    $selectedStore = Get-StoreByIdOrDefault -namespace $namespace -storeId $StoreId
    if (-not $selectedStore) { throw "No se pudo resolver el store por defecto." }
    $resolvedStoreId = $null
    try { $resolvedStoreId = $selectedStore.StoreID } catch {}

    $yearCounts = @{}
    $yearSizes = @{}
    $monthCounts = @{}
    $monthSizes = @{}
    $folderStats = @{}
    $totalMailboxItems = 0

    $rootFolders = Get-SubFolders-Safe -parentFolder $selectedStore.GetRootFolder()
    $foldersToAnalyze = @()
    if ($rootFolders) {
        foreach ($f in $rootFolders) {
            if ($excludedFolderNames -contains $f.Name) { continue }
            $foldersToAnalyze += $f
        }
    }

    $totalFolders = $foldersToAnalyze.Count
    $processedFolders = 0

    foreach ($f in $foldersToAnalyze) {
        $processedFolders++
        $percent = 0
        if ($totalFolders -gt 0) {
            $percent = [int][math]::Round(($processedFolders / $totalFolders) * 100)
            if ($percent -gt 99 -and $processedFolders -lt $totalFolders) { $percent = 99 }
        }

        Write-Progress -Activity "Análisis de buzón" -Status "Carpeta $processedFolders/$($totalFolders): $($f.Name)" -PercentComplete $percent
        Write-Host "Progreso análisis: $percent% ($processedFolders/$totalFolders) - $($f.Name)" -ForegroundColor DarkGray
        Write-Host "Analizando carpeta: $($f.Name)..." -ForegroundColor Gray
        $folderCount = Get-YearDistributionRecursive -folder $f -yearCounts $yearCounts -yearSizes $yearSizes -monthCounts $monthCounts -monthSizes $monthSizes -folderStats $folderStats -yearFilter $BackupYear -monthsFilter $BackupMonthsList
        $totalMailboxItems += $folderCount
    }

    Write-Progress -Activity "Análisis de buzón" -Status "Completado" -PercentComplete 100 -Completed

    $years = $yearCounts.Keys | Sort-Object -Descending
    $dist = @()
    foreach ($y in $years) {
        $monthsList = @()
        # Sort months 1..12
        for ($m = 1; $m -le 12; $m++) {
            $key = "$y-$m"
            if ($monthCounts.ContainsKey($key)) {
                $monthsList += [pscustomobject]@{
                    month = $m
                    count = [int]$monthCounts[$key]
                    sizeBytes = [long]$monthSizes[$key]
                    sizeHuman = (Format-Bytes -bytes $monthSizes[$key])
                }
            }
        }

        $breakdown = @()
        foreach ($fp in $folderStats.Keys) {
            if ($folderStats[$fp].ContainsKey($y)) {
                $breakdown += [pscustomobject]@{
                    path = $fp
                    count = [int]$folderStats[$fp][$y]
                }
            }
        }
        $breakdown = $breakdown | Sort-Object count -Descending

        $dist += [pscustomobject]@{
            year = [int]$y
            count = [int]$yearCounts[$y]
            sizeBytes = [long]$yearSizes[$y]
            sizeHuman = (Format-Bytes -bytes $yearSizes[$y])
            months = @($monthsList)
            folderBreakdown = @($breakdown)
        }
    }

    Write-Output (@{
        type = "analysis"
        store = @{
            displayName = $selectedStore.DisplayName
            storeId = $resolvedStoreId
        }
        totalItems = [int]$totalMailboxItems
        distribution = @($dist)
    } | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue)
    exit 0
}

if ($Command -eq "run-backup") {
    if (-not $BackupYear) {
        Write-Error "Debe proporcionar -BackupYear para ejecutar run-backup."
        exit 1
    }
    $selectedStore = Get-StoreByIdOrDefault -namespace $namespace -storeId $StoreId
    if ($selectedStore) {
        $sourceFolders = $selectedStore.GetRootFolder().Folders
    }
}

# Interactive Mode if BackupYear is not provided
if ($Command -eq "interactive" -and -not $BackupYear) {
    Clear-Host
    Write-Host "=== MODO INTERACTIVO DE BACKUP DE OUTLOOK ===" -ForegroundColor Cyan
    Write-Host "---------------------------------------------" -ForegroundColor Gray
    
    # 1. List available stores
    $connectedStores = Get-ConnectedOutlookStores -namespace $namespace -IncludeRawStore
    $storeList = @()
    $i = 1
    
    Write-Host "Buzones y Archivos de Datos disponibles:" -ForegroundColor Yellow
    foreach ($entry in $connectedStores) {
        $store = $entry.store
        $storeInfo = $entry.info
        # Show ALL stores to ensure visibility of Online archives, Shared Mailboxes, etc.
        $pathDisplay = "Modo Online / Sin ruta local"
        $extraInfo = ""
        
        if ($storeInfo.filePath) {
            $pathDisplay = $storeInfo.filePath
        }
        
        # Try to identify Exchange type for better context
        try {
            switch ($storeInfo.exchangeStoreType) {
                1 { $extraInfo = "[Exchange Principal]" }
                2 { $extraInfo = "[Exchange Delegado]" }
                3 { $extraInfo = "[Carpetas Públicas]" }
            }
        } catch {}
        
        # Display everything
        Write-Host "[$i] $($storeInfo.displayName) $extraInfo - ($pathDisplay)"
        $storeList += $store
        $i++
    }
    
    if ($storeList.Count -eq 0) {
        Write-Warning "No se encontraron buzones Exchange ni archivos .OST conectados."
        exit
    }
    
    Write-Host ""
    $selection = Read-Host "Seleccione el número del buzón a respaldar"
    
    # Cast to int to ensure correct numerical comparison
    try {
        $selectionInt = [int]$selection
    } catch {
        $selectionInt = 0
    }

    if ($selectionInt -gt 0 -and $selectionInt -le $storeList.Count) {
        $selectedStore = $storeList[$selectionInt-1]
        Write-Host "Seleccionado: $($selectedStore.DisplayName)" -ForegroundColor Green
        
        # --- Count Total Items in Mailbox & Group by Year ---
        Write-Host "Analizando estructura del buzón y distribución por años... (Esto puede tardar según el tamaño)" -ForegroundColor Cyan
        $totalMailboxItems = 0
        $yearCounts = @{} # Hashtable to store counts per year
        $yearSizes = @{}  # Hashtable to store size in bytes per year
        
        # Helper function to format bytes
        function Format-Bytes {
            param ($bytes)
            if ($bytes -ge 1GB) {
                return "{0:N2} GB" -f ($bytes / 1GB)
            } elseif ($bytes -ge 1MB) {
                return "{0:N2} MB" -f ($bytes / 1MB)
            } elseif ($bytes -ge 1KB) {
                return "{0:N2} KB" -f ($bytes / 1KB)
            } else {
                return "$bytes Bytes"
            }
        }

        # Helper function for total count and year grouping (Optimized with GetTable)
        function Analyze-ItemsRecursive {
            param ($folder)
            $cnt = 0
            
            try {
                if ($folder.Items.Count -gt 0) {
                    try {
                        # Optimization: Use GetTable instead of iterating Items (100x faster)
                        # We only need ReceivedTime (or CreationTime) and Size
                        $filter = "" # No filter, get all
                        $table = $folder.GetTable($filter)
                        
                        # Remove default columns to save memory
                        $table.Columns.RemoveAll()
                        # Add only needed columns
                        $table.Columns.Add("ReceivedTime") | Out-Null
                        $table.Columns.Add("CreationTime") | Out-Null
                        $table.Columns.Add("Size") | Out-Null
                        
                        # Fetch all rows
                        while (-not $table.EndOfTable) {
                            $rows = $table.GetNextRows(100) # Batch size 100
                            foreach ($row in $rows) {
                                $date = $row["ReceivedTime"]
                                if (-not $date) { $date = $row["CreationTime"] }
                                
                                if ($date) {
                                    $y = $date.Year
                                    $s = $row["Size"]
                                    
                                    if ($yearCounts.ContainsKey($y)) {
                                        $yearCounts[$y]++
                                        $yearSizes[$y] += $s
                                    } else {
                                        $yearCounts[$y] = 1
                                        $yearSizes[$y] = $s
                                    }
                                    $cnt++
                                }
                            }
                        }
                    } catch {
                        # Fallback to slow iteration if GetTable fails (unlikely on modern Outlook)
                        foreach ($item in $folder.Items) {
                             $date = $item.ReceivedTime
                             if (-not $date) { $date = $item.CreationTime }
                             if ($date) {
                                $y = $date.Year
                                $s = $item.Size
                                if ($yearCounts.ContainsKey($y)) { $yearCounts[$y]++; $yearSizes[$y] += $s } 
                                else { $yearCounts[$y] = 1; $yearSizes[$y] = $s }
                                $cnt++
                             }
                        }
                    }
                }
            } catch {}
            
            foreach ($sub in $folder.Folders) {
                $cnt += Analyze-ItemsRecursive -folder $sub
            }
            return $cnt
        }

        function Get-RomanNumeral {
            param ([int]$number)
            if ($number -lt 1 -or $number -gt 3999) { return $number.ToString() }
            $dict = @{ 1000="M"; 900="CM"; 500="D"; 400="CD"; 100="C"; 90="XC"; 50="L"; 40="XL"; 10="X"; 9="IX"; 5="V"; 4="IV"; 1="I" }
            $keys = $dict.Keys | Sort-Object -Descending
            $result = ""
            foreach ($key in $keys) {
                while ($number -ge $key) {
                    $result += $dict[$key]
                    $number -= $key
                }
            }
            return $result
        }

        # Calculate total and display per-folder breakdown
        foreach ($f in $selectedStore.GetRootFolder().Folders) {
            # Skip typically empty or system folders to save time if needed, 
            # but for accuracy we check everything.
            $folderCount = Analyze-ItemsRecursive -folder $f
            if ($folderCount -gt 0) {
                Write-Host " - $($f.Name): " -NoNewline
                Write-Host "$folderCount elementos" -ForegroundColor Yellow
                $totalMailboxItems += $folderCount
            }
        }
        
        Write-Host "---------------------------------------------" -ForegroundColor Gray
        Write-Host "Total de elementos en el buzón: $totalMailboxItems" -ForegroundColor White -BackgroundColor DarkBlue
        
        Write-Host ""
        Write-Host "Distribución por Años:" -ForegroundColor Cyan
        # Sort years descending
        $sortedYears = $yearCounts.Keys | Sort-Object -Descending
        foreach ($y in $sortedYears) {
            $formattedSize = Format-Bytes -bytes $yearSizes[$y]
            Write-Host " - Año $($y) ($formattedSize): " -NoNewline
            Write-Host "$($yearCounts[$y]) elementos" -ForegroundColor Green
        }
        Write-Host "---------------------------------------------" -ForegroundColor Gray
        Write-Host ""
        # ------------------------------------

        # In interactive mode, we backup ALL top-level folders of the selected store
        # to ensure we capture everything the user likely wants in a full backup.
        $sourceFolders = $selectedStore.GetRootFolder().Folders
        
    } else {
        Write-Error "Selección inválida."
        exit 1
    }
    
    Write-Host ""
    # 2. Ask for Year
    do {
        $yearInput = Read-Host "Ingrese el año de backup (ej. 2023)"
        if ($yearInput -match "^\d{4}$") {
            $BackupYear = [int]$yearInput
        } else {
            Write-Host "Por favor ingrese un año válido de 4 dígitos." -ForegroundColor Red
        }
    } until ($BackupYear)
    
    # 3. Ask for Path
    Write-Host ""
    $pathInput = Read-Host "Ingrese ruta de destino (Presione Enter para usar Default: C:\Backups)"
    if (-not [string]::IsNullOrWhiteSpace($pathInput)) {
        $TargetBackupPath = $pathInput
    }
    
    Write-Host ""
    $baseNameForDefault = "Outlook"
    if ($selectedStore) {
        $baseNameForDefault = $selectedStore.DisplayName
    }
    $safeBaseNameForDefault = $baseNameForDefault -replace '[\\/:*?"<>|]', '_'
    $defaultPstName = "Backup_${safeBaseNameForDefault}_$BackupYear"
    $pstInput = Read-Host "Ingrese nombre del PST (Default: $defaultPstName)"
    if (-not [string]::IsNullOrWhiteSpace($pstInput)) {
        $PstName = $pstInput
    }
    
    Write-Host ""
    # 4. Ask for Action (Copy or Move)
    Write-Host "Seleccione el tipo de operacion:" -ForegroundColor Cyan
    Write-Host "[1] COPIAR (Backup - Mantiene los correos originales en el buzon)" -ForegroundColor Green
    Write-Host "[2] MOVER  (Archivar - Mueve los correos al PST y libera espacio)" -ForegroundColor Yellow
    
    while ($true) {
        $actionInput = Read-Host "Opcion [1/2] (Default: 1)"
        if ([string]::IsNullOrWhiteSpace($actionInput) -or $actionInput -eq "1") {
            $BackupAction = "Copy"
            break
        }
        if ($actionInput -eq "2") {
            $BackupAction = "Move"
            Write-Warning "ATENCION: Ha seleccionado MOVER. Los correos seran eliminados de su buzon principal tras copiarse al PST."
            break
        }
        Write-Host "Opcion invalida." -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Iniciando configuración del backup..." -ForegroundColor Cyan
}

# Define backup path name logic
$baseName = "Outlook"
if ($selectedStore) {
    $baseName = $selectedStore.DisplayName
}

# 1. OPTIMIZACIÓN NOMBRE: Formato "BCK [AÑO] [nombre del buzon antes del arroba]"
# Extraer parte antes del arroba si existe
if ($baseName -match "^([^@]+)@") {
    $namePart = $matches[1]
} else {
    $namePart = $baseName
}

# Limpiar caracteres no alfanuméricos y espacios extra
$safeNamePart = $namePart -replace '[^a-zA-Z0-9 ]', ''
$safeNamePart = $safeNamePart.Trim() -replace '\s+', ' '

# Convertir a MAYUSCULAS
$safeNamePart = $safeNamePart.ToUpper()

if (-not $PstName) {
    # Formato solicitado: BCK [AÑO] [NOMBRE]
    $PstName = "BCK $BackupYear $safeNamePart"
}

# Limpieza final del PstName por si viene por parametro (asegurar caracteres validos)
$PstName = ($PstName -replace '[^a-zA-Z0-9 _-]', '').Trim()

if ($PstName.EndsWith(".pst", [System.StringComparison]::OrdinalIgnoreCase)) {
    $PstName = $PstName.Substring(0, $PstName.Length - 4)
}
$backupFileName = "$PstName.pst"
$pstFilePath = Join-Path $TargetBackupPath $backupFileName

# Ensure the backup directory exists
Write-Host "Verificando directorio destino: $TargetBackupPath ..." -ForegroundColor Cyan
if (-not (Test-Path $TargetBackupPath)) {
    try {
        New-Item -ItemType Directory -Force -Path $TargetBackupPath | Out-Null
        Write-Host "Directorio de backup creado: $TargetBackupPath" -ForegroundColor Green
    } catch {
        Write-Error "No se pudo crear el directorio de backup: $($_.Exception.Message)"
        exit 1
    }
}

# El nombre visual en Outlook debe coincidir con el nombre del archivo (sin extensión)
$pstStoreName = $PstName
$pstRootFolder = $null
$pstStore = $null

function Normalize-PathValue {
    param([string]$pathValue)
    if ([string]::IsNullOrWhiteSpace($pathValue)) { return $null }
    return $pathValue.Trim().ToLowerInvariant()
}

function Find-PstStoreByPath {
    param($stores, [string]$targetPath)
    $targetNorm = Normalize-PathValue $targetPath
    if (-not $targetNorm) { return $null }
    foreach ($s in $stores) {
        try {
            $fp = $s.FilePath
        } catch {
            continue
        }
        $fpNorm = Normalize-PathValue $fp
        if ($fpNorm -and $fpNorm -eq $targetNorm) { return $s }
    }
    return $null
}

function Set-PstDisplayName {
    param($store, [string]$desiredName)
    if (-not $store -or [string]::IsNullOrWhiteSpace($desiredName)) { return $false }
    $currentName = $null
    try { $currentName = $store.DisplayName } catch {}
    if ($currentName -eq $desiredName) { return $true }
    try {
        $store.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3001001F", $desiredName)
    } catch {
        Write-Warning "No se pudo establecer el nombre del PST: $($_.Exception.Message)"
        return $false
    }
    Start-Sleep -Milliseconds 200
    $updatedName = $null
    try { $updatedName = $store.DisplayName } catch {}
    if ($updatedName -ne $desiredName) {
        Write-Warning "El nombre del PST no coincide en Outlook. Actual: '$updatedName' / Esperado: '$desiredName'"
        return $false
    }
    return $true
}

function Get-OutlookPSTPolicy {
    $versions = @("16.0","15.0","14.0")
    $result = @{}
    foreach ($v in $versions) {
        foreach ($hive in @("HKCU","HKLM")) {
            $path = "${hive}:\Software\Policies\Microsoft\Office\${v}\Outlook"
            try {
                if (Test-Path $path) {
                    $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
                    if ($props) {
                        foreach ($name in @("DisablePST","PSTDisableGrow","DisableCrossAccountCopy")) {
                            if ($props.PSObject.Properties[$name]) { $result[$name] = $props.$name }
                        }
                    }
                }
            } catch {}
        }
    }
    return $result
}

Write-Host "Comprobando políticas de Outlook..." -ForegroundColor Gray
$pstPolicy = Get-OutlookPSTPolicy
if ($pstPolicy.ContainsKey("DisablePST") -and $pstPolicy["DisablePST"] -eq 1) {
    Write-Error "La política de Outlook impide crear/montar PST (DisablePST=1)."
    exit 1
}
if ($pstPolicy.ContainsKey("PSTDisableGrow") -and $pstPolicy["PSTDisableGrow"] -eq 1) {
    Write-Warning "La política de Outlook impide el crecimiento de PST existentes (PSTDisableGrow=1)."
}
if ($pstPolicy.ContainsKey("DisableCrossAccountCopy") -and $pstPolicy["DisableCrossAccountCopy"] -eq 1) {
    Write-Warning "La política de Outlook impide copiar elementos entre cuentas (DisableCrossAccountCopy=1)."
}

Write-Host "Gestionando archivo PST: $backupFileName ..." -ForegroundColor Cyan

# Check if PST already exists and if it's open in Outlook
Write-Host "Verificando si el PST ya está montado en Outlook..." -ForegroundColor Gray
$existingStoreInOutlook = Find-PstStoreByPath -stores $namespace.Stores -targetPath $pstFilePath

if ($existingStoreInOutlook) {
    Write-Host "El archivo PST ya está conectado. Usando existente." -ForegroundColor Yellow
    $pstRootFolder = $existingStoreInOutlook.GetRootFolder()
    $pstStore = $existingStoreInOutlook
    $nameOk = Set-PstDisplayName -store $pstStore -desiredName $pstStoreName
    if (-not $nameOk) {
        Write-Warning "No se pudo confirmar el nombre del PST en Outlook."
    }
} else {
    Write-Host "El PST no está montado. Verificando existencia en disco..." -ForegroundColor Gray
    # If PST file exists on disk but not in Outlook, add it.
    try {
        if (Test-Path $pstFilePath) {
             Write-Host "Archivo encontrado en disco. Montando..." -ForegroundColor Cyan
             $namespace.AddStore($pstFilePath)
             Write-Host "Montaje completado." -ForegroundColor Green
        } else {
            Write-Host "Archivo no existe. Creando nuevo PST..." -ForegroundColor Cyan
            Write-Progress -Activity "Gestión de PST" -Status "Creando archivo PST (esto puede tardar unos segundos)..." -PercentComplete 0
            try {
                Write-Host "Intentando AddStore estándar..." -ForegroundColor Gray
                try {
                    $namespace.AddStore($pstFilePath)
                    Write-Host "AddStore completado." -ForegroundColor Green
                    Start-Sleep -Seconds 1
                } catch {
                    $err = $_.Exception.Message
                    Write-Host "AddStore estándar falló: $err" -ForegroundColor Yellow
                    Write-Host "Intentando AddStoreEx (Unicode)..." -ForegroundColor Gray
                    Write-Progress -Activity "Gestión de PST" -Status "Reintentando con AddStoreEx..." -PercentComplete 0
                    # AddStoreEx solo acepta (Path, Type). El nombre se configura despues via PropertyAccessor.
                    $namespace.AddStoreEx($pstFilePath, 2)
                    Write-Host "AddStoreEx completado." -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
            } catch {
                Write-Error "No se pudo crear el nuevo PST: $($_.Exception.Message)"
                exit 1
            } finally {
                Write-Progress -Activity "Gestión de PST" -Status "PST Creado" -Completed
            }
        }
        
        # Retrieve the newly added store
        Write-Host "Recuperando referencia al nuevo Store..." -ForegroundColor Gray
        $pstStore = Find-PstStoreByPath -stores $namespace.Stores -targetPath $pstFilePath

        if ($pstStore) {
            $pstRootFolder = $pstStore.GetRootFolder()
            $nameOk = Set-PstDisplayName -store $pstStore -desiredName $pstStoreName
            if (-not $nameOk) {
                Write-Warning "No se pudo confirmar el nombre del PST en Outlook."
            }
            # Confirm file presence on disk
            if (-not (Test-Path $pstFilePath)) {
                Write-Warning "El archivo PST aún no aparece en disco: $pstFilePath. Verifique permisos o políticas."
            }
        } else {
            throw "No se pudo recuperar el store recién añadido. Verifique si el archivo se creó en $pstFilePath"
        }
    } catch {
        Write-Error "Error crítico al gestionar el archivo PST: $($_.Exception.Message)"
        exit 1
    }
}

if (-not $pstRootFolder) {
    Write-Host "No se pudo obtener la carpeta raíz del archivo PST. Abortando."
    exit 1
}

# Define the date range for the specified year
$startDate = Get-Date -Year ([int]$BackupYear) -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
$endDate = Get-Date -Year (([int]$BackupYear) + 1) -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0

# If sourceFolders is not set (Non-interactive mode or default logic), use default folders
if (-not $sourceFolders) {
    # Outlook folder constants
    $olFolderInbox = 6
    $olFolderSentMail = 5
    $olFolderDrafts = 16
    $olFolderDeletedItems = 3

    $sourceFolders = @(
        $namespace.GetDefaultFolder($olFolderInbox),
        $namespace.GetDefaultFolder($olFolderSentMail),
        $namespace.GetDefaultFolder($olFolderDrafts),
        $namespace.GetDefaultFolder($olFolderDeletedItems)
    )
}

Write-Host "Iniciando el backup de correos del año $BackupYear..."
Write-Host "Modo: $BackupAction"
Write-Host "Destino: $pstFilePath"

# Global counters and PST management
$script:totalItemsToProcess = 0
$script:processedItems = 0
$script:currentPstPart = 1
$script:currentPstSize = 0
$script:maxPstSize = $null # Null means no limit (or default single PST)
$script:basePstName = ""

# --- Recursive Functions ---

function Get-RecursiveItemCount {
    param (
        $folder,
        $filter,
        [int[]]$monthsToInclude
    )
    
    $count = 0
    
    # Exclusion Check
    if ($excludedFolderNames -contains $folder.Name) {
        return 0
    }

    # Count items in current folder
    try {
        if ($folder.Items.Count -gt 0) {
            $useManual = $false

            try {
                $table = $folder.GetTable($filter)
                $table.Columns.RemoveAll()
                $table.Columns.Add("ReceivedTime") | Out-Null
                $table.Columns.Add("CreationTime") | Out-Null
                
                while (-not $table.EndOfTable) {
                   $rows = $table.GetNextRows(500)
                   
                   foreach ($row in $rows) {
                       # Month filtering check
                       if ($monthsToInclude) {
                           $d = $row["ReceivedTime"]
                           if (-not $d) { $d = $row["CreationTime"] }
                           if ($d -and ($d.Month -notin $monthsToInclude)) {
                               continue
                           }
                       }
                       $count++
                   }
                }
            } catch {
                try {
                    $itemsFiltered = $folder.Items.Restrict($filter)
                    if ($monthsToInclude) {
                        foreach ($item in $itemsFiltered) {
                            $d = $item.ReceivedTime
                            if (-not $d) { $d = $item.CreationTime }
                            if ($d -and ($d.Month -in $monthsToInclude)) {
                                $count++
                            }
                        }
                    } else {
                        $count += $itemsFiltered.Count
                    }
                } catch {
                    $useManual = $true
                }
            }

            if ($useManual) {
                 # Fallback manual puro
                 $startD = Get-Date -Year ([int]$BackupYear) -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                 $endD = Get-Date -Year (([int]$BackupYear) + 1) -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                 
                 foreach ($item in $folder.Items) {
                     $d = $item.ReceivedTime
                     if (-not $d) { $d = $item.CreationTime }
                     
                     if ($d -and $d -ge $startD -and $d -lt $endD) {
                        if ($monthsToInclude -and ($d.Month -notin $monthsToInclude)) { continue }
                        $count++
                     }
                 }
            }
        }
    } catch {
        # Some folders might error out on Restrict or Access
        # Just ignore and continue
    }
    
    # Recursively count subfolders
    $subFolders = Get-SubFolders-Safe -parentFolder $folder
    if ($subFolders) {
        foreach ($subFolder in $subFolders) {
            $count += Get-RecursiveItemCount -folder $subFolder -filter $filter -monthsToInclude $monthsToInclude
        }
    }
    
    return $count
}

function Switch-PST {
    param (
        $sourceFolderPath
    )
    
    Write-Host "`n[INFO] Límite de tamaño alcanzado ($([math]::Round($script:currentPstSize / 1GB, 2)) GB). Rotando PST..." -ForegroundColor Yellow
    
    # 1. Close current PST
    if ($global:pstStore) {
        try {
            $namespace.RemoveStore($global:pstStore.GetRootFolder())
        } catch {
            Write-Warning "Error al cerrar PST actual: $($_.Exception.Message)"
        }
    }
    
    # 2. Increment Part
    $script:currentPstPart++
    $script:currentPstSize = 0
    
    # 3. Create new PST Name
    # Format: "BCK [Year] [Roman] [MailboxName]"
    # Wait, user requested: "BCK 2024 I [Nombre del buzon]"
    # My $script:basePstName should hold "[Nombre del buzon]" (cleaned)
    $roman = Get-RomanNumeral $script:currentPstPart
    $newFileName = "BCK $($BackupYear) $roman $script:basePstName.pst"
    $newPstPath = Join-Path $TargetBackupPath $newFileName
    
    Write-Host "[INFO] Creando nuevo PST: $newFileName" -ForegroundColor Cyan
    
    # 4. Create/Mount new PST
    try {
        $namespace.AddStore($newPstPath)
    } catch {
        try {
             $namespace.AddStoreEx($newPstPath, 2)
        } catch {
             throw "No se pudo crear el siguiente archivo PST: $newPstPath"
        }
    }
    
    # 5. Get Reference
    $global:pstStore = $null
    $targetNorm = Normalize-PathValue $newPstPath
    foreach ($s in $namespace.Stores) {
        try {
            if ((Normalize-PathValue $s.FilePath) -eq $targetNorm) { $global:pstStore = $s; break }
        } catch {}
    }
    
    if (-not $global:pstStore) { throw "No se pudo recuperar el nuevo PST." }
    
    $global:pstRootFolder = $global:pstStore.GetRootFolder()
    try { $global:pstStore.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3001001F", "Backup Part $roman") } catch {}
    
    # 6. Recreate Folder Path in New PST
    # We need to find the equivalent of $sourceFolderPath in the new PST
    # $sourceFolderPath format: "\\user@domain.com\Inbox\Subfolder"
    # We need to strip the store name and traverse
    
    # Simple traversal based on folder names
    # Assuming $sourceFolderPath is the full path. We need to parse it.
    # Actually, simpler: We are deep in recursion. The 'Process-RecursiveItems' function
    # has a $targetPstFolder parameter which is now STALE.
    # We need to return the NEW target folder for the current source folder.
    
    # To do this, we need to know the relative path of the source folder from the source root.
    # But we don't have that easily. 
    # Alternative: The $sourceFolderPath passed to this function is the FolderPath property.
    # Example: \\Mailbox\Inbox\Project
    # PST Root: \\Personal Folders
    # We want: \\Personal Folders\Inbox\Project
    
    $pathParts = $sourceFolderPath -split '\\'
    # First part is empty (root), second is Store Name. Remove them.
    $relParts = $pathParts | Select-Object -Skip 2
    
    $current = $global:pstRootFolder
    foreach ($part in $relParts) {
        if (-not [string]::IsNullOrWhiteSpace($part)) {
            try {
                $found = $current.Folders | Where-Object { $_.Name -eq $part }
                if (-not $found) {
                    $found = $current.Folders.Add($part)
                }
                $current = $found
            } catch {
                # Fallback for folder creation errors
                return $global:pstRootFolder # Should not happen ideally
            }
        }
    }
    
    return $current
}

function Process-RecursiveItems {
    param (
        $sourceFolder,
        $targetPstFolder,
        $filter,
        $action,
        [int[]]$monthsToInclude
    )

    Write-Host "Explorando: $($sourceFolder.FolderPath)"

    # Exclusion Check (Safety)
    if ($excludedFolderNames -contains $sourceFolder.Name) {
        Write-Host "  [SKIP] Carpeta excluida: $($sourceFolder.Name)" -ForegroundColor DarkGray
        return
    }

    # 1. Process items in current folder
    try {
        if ($sourceFolder.Items.Count -gt 0) {
            $itemsArray = @()
            $useManual = $false

            try {
                # Intentar filtrado rápido (Restrict)
                # El filtro ya incluye logica OR (Received/Creation) gracias a la corrección anterior
                $itemsToProcess = $sourceFolder.Items.Restrict($filter)
                foreach ($item in $itemsToProcess) { 
                    # Check Month if needed
                    if ($monthsToInclude) {
                        $d = $item.ReceivedTime
                        if (-not $d) { $d = $item.CreationTime }
                        if ($d -and ($d.Month -notin $monthsToInclude)) {
                            continue
                        }
                    }
                    $itemsArray += $item 
                }
            } catch {
                $useManual = $true
                Write-Host "  [WARN] Filtrado rápido falló en '$($sourceFolder.Name)'. Usando modo manual (lento)..." -ForegroundColor Yellow
            }

            if ($useManual) {
                # Fallback manual: Iterar todo y chequear fechas
                # Recalculamos fechas limite basadas en BackupYear global
                $startD = Get-Date -Year $BackupYear -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                $endD = Get-Date -Year (([int]$BackupYear) + 1) -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0

                foreach ($item in $sourceFolder.Items) {
                    $d = $item.ReceivedTime
                    if (-not $d) { $d = $item.CreationTime }
                    
                    if ($d -and $d -ge $startD -and $d -lt $endD) {
                        if ($monthsToInclude -and ($d.Month -notin $monthsToInclude)) { continue }
                        $itemsArray += $item
                    }
                }
            }
            
            $totalInFolder = $itemsArray.Count
            if ($totalInFolder -gt 0) {
                Write-Host "  > Procesando $totalInFolder elementos ($action)..." -ForegroundColor Gray
                
                # We need to update targetPstFolder if it changed during recursion
                # But here we are iterating items. If Switch-PST happens, 
                # we need to update our local $targetPstFolder variable.
                
                foreach ($item in $itemsArray) {
                    try {
                        # Check Size Limit for Split Mode
                        if ($script:maxPstSize -and $script:currentPstSize -ge $script:maxPstSize) {
                            $targetPstFolder = Switch-PST -sourceFolderPath $sourceFolder.FolderPath
                        }
                    
                        $success = $false
                        $itemSize = $item.Size
                        
                        if ($action -eq "Move" -or $action -eq "move") {
                            # MOVE logic
                            $item.Move($targetPstFolder) | Out-Null
                            $success = $true
                        } else {
                            # COPY logic
                            $copy = $item.Copy()
                            if ($copy) {
                                $copy.Move($targetPstFolder) | Out-Null
                                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($copy) | Out-Null
                                $success = $true
                            }
                        }

                        if ($success) {
                            $script:processedItems++
                            $script:currentPstSize += $itemSize
                        }
                        
                        if ($script:totalItemsToProcess -gt 0) {
                            $percentComplete = [math]::Min(100, [math]::Round(($script:processedItems / $script:totalItemsToProcess) * 100))
                        } else {
                            $percentComplete = 0
                        }

                        $actMsg = if ($action -eq "Move" -or $action -eq "move") { "Moviendo" } else { "Copiando" }
                        Write-Progress -Activity "Realizando backup de correos de Outlook $BackupYear" `
                                       -Status "$actMsg ($($script:processedItems)/$($script:totalItemsToProcess)): $($item.Subject)" `
                                       -PercentComplete $percentComplete
                    } catch {
                        # Ignore harmless errors like "Object moved" or minor property access issues
                        Write-Host "  ! Error leve en item: $($_.Exception.Message)" -ForegroundColor DarkGray
                    }
                }
            }
        }
    } catch {
        # Catch-all for folder access errors (e.g. unknown properties in some item types)
        Write-Host "  ! Error acceso carpeta $($sourceFolder.Name): $($_.Exception.Message)" -ForegroundColor DarkGray
    }

    # 2. Process subfolders
    foreach ($subFolder in $sourceFolder.Folders) {
        try {
            # Crucial: Re-evaluate targetPstFolder in case it changed in parent scope (unlikely as we pass by value)
            # Actually, if Switch-PST happened above, $targetPstFolder is updated for THIS scope.
            # But we need to ensure the subfolder is created in the CURRENT active PST.
            
            # Since Switch-PST updates global state ($global:pstRootFolder), we should derive target from there?
            # Or just rely on $targetPstFolder being updated in the item loop?
            # If the folder was empty of items, Switch-PST wouldn't have run. 
            # So we use the current $targetPstFolder (which might be the old one).
            # This is fine. Split only happens when we hit size limit WRITING items.
            
            # However, if we just switched PST in the item loop, $targetPstFolder is NEW.
            # So subfolders will be created in the NEW PST. Correct.
            
            # 2. OPTIMIZACIÓN CARPETAS VACIAS (Recursiva)
            $subItemsCount = Get-RecursiveItemCount -folder $subFolder -filter $filter -monthsToInclude $monthsToInclude
            
            if ($subItemsCount -gt 0) {
                # Get or create corresponding folder in PST
                $targetSubFolder = $targetPstFolder.Folders | Where-Object {$_.Name -eq $subFolder.Name}
                if (-not $targetSubFolder) {
                    $targetSubFolder = $targetPstFolder.Folders.Add($subFolder.Name)
                }
                
                # Recursive call
                Process-RecursiveItems -sourceFolder $subFolder -targetPstFolder $targetSubFolder -filter $filter -action $action -monthsToInclude $monthsToInclude
            }
            
        } catch {
            Write-Host "  Error al procesar subcarpeta $($subFolder.Name): $($_.Exception.Message)"
        }
    }
}

# --- Execution ---

# Filter string
$startStr = $startDate.ToString('MM/dd/yyyy HH:mm tt')
$endStr = $endDate.ToString('MM/dd/yyyy HH:mm tt')
$filter = "([ReceivedTime] >= '$startStr' AND [ReceivedTime] < '$endStr') OR ([ReceivedTime] IS NULL AND [CreationTime] >= '$startStr' AND [CreationTime] < '$endStr')"

# Optimization: If we already have the count from interactive mode analysis, use it.
if ($yearCounts -and $yearCounts.ContainsKey($BackupYear) -and (-not $BackupMonthsList)) {
    $script:totalItemsToProcess = $yearCounts[$BackupYear]
    Write-Host "Usando conteo pre-calculado para el año $($BackupYear): $script:totalItemsToProcess elementos." -ForegroundColor Cyan
} else {
    # First pass: Count total items (Recursive) - Only if we don't have the count yet
    Write-Host "Contando elementos para el backup (esto puede tardar unos momentos)..."
    foreach ($sourceFolder in $sourceFolders) {
        if ($excludedFolderNames -contains $sourceFolder.Name) { continue }
        $script:totalItemsToProcess += Get-RecursiveItemCount -folder $sourceFolder -filter $filter -monthsToInclude $BackupMonthsList
    }
}

Write-Host "Total de correos encontrados para procesar: $script:totalItemsToProcess"

# Set Base Pst Name for Split Logic
if ($PstName) {
    # If user provided a name via parameter or input, use it as base
    $script:basePstName = $PstName -replace '^Backup_\d{4}_', '' # Clean if user re-used pattern
    $script:basePstName = $script:basePstName -replace '^BCK \d{4} [IVXLCDM]+ ', '' # Clean split pattern
} else {
    $script:basePstName = $safeNamePart
}

# Check Size and Enable Split Mode if > 40GB
# We use $yearSizes[$BackupYear] if available (Interactive mode)
# If not available (Non-interactive), we might need to count size first.
# For now, we assume interactive flow populated $yearSizes.
# If running non-interactive, $yearSizes is empty. We should probably calculate it in First Pass if critical.
# However, user prompt implies this is for the interactive flow context.
if ($yearSizes -and $yearSizes.ContainsKey($BackupYear)) {
    $totalSizeBytes = $yearSizes[$BackupYear]
    if ($totalSizeBytes -gt 40GB) {
        Write-Host "Detectado tamaño grande ($([math]::Round($totalSizeBytes/1GB, 2)) GB). Activando modo Split (20GB/PST)." -ForegroundColor Yellow
        $script:maxPstSize = 20GB
        
        # Override the initial PST name to match the split convention
        # First file: "BCK [Year] I [Name]"
        $roman = Get-RomanNumeral 1
        $initialSplitName = "BCK $($BackupYear) $roman $script:basePstName.pst"
        
        # We need to rename/remount the PST created earlier or close and create new.
        # The script created $backupFileName earlier.
        # Let's close it and switch to the split name.
        
        if ($pstStore) {
             try { $namespace.RemoveStore($pstStore.GetRootFolder()) } catch {}
             # Delete the initial generic file if empty/unused or rename it?
             # Better to just start fresh or rename.
             # Renaming file is tricky while Outlook might have lock.
             # Let's just create the new one and ignore the old empty one for now.
        }
        
        $backupFileName = $initialSplitName
        $pstFilePath = Join-Path $TargetBackupPath $backupFileName
        
        # Mount the new first Split PST
        try {
             $namespace.AddStore($pstFilePath)
        } catch {
             $namespace.AddStoreEx($pstFilePath, 2)
        }
        
        # Update Reference
        $pstStore = $null
        $targetNorm = Normalize-PathValue $pstFilePath
        foreach ($s in $namespace.Stores) {
            try {
                if ((Normalize-PathValue $s.FilePath) -eq $targetNorm) { $pstStore = $s; break }
            } catch {}
        }
        if ($pstStore) {
             $pstRootFolder = $pstStore.GetRootFolder()
             $global:pstStore = $pstStore # Sync global
             $global:pstRootFolder = $pstRootFolder
             try { $pstStore.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3001001F", "Backup Part I") } catch {}
        }
    }
}

# Second pass: Process items (Recursive) with Robust Connection Handling

# Pre-fetch EntryIDs to avoid iterator invalidation if connection drops
$folderIds = @()
foreach ($f in $sourceFolders) {
    try { $folderIds += $f.EntryID } catch {}
}

foreach ($fid in $folderIds) {
    $retryCount = 0
    $maxRetries = 3
    $sourceFolder = $null
    
    # Try to get folder object with retries
    while ($retryCount -lt $maxRetries) {
        try {
            $sourceFolder = $namespace.GetFolderFromID($fid)
            break
        } catch {
            $retryCount++
            Write-Warning "Perdida de conexión al acceder a carpeta (Intento $retryCount/$maxRetries): $($_.Exception.Message). Esperando..."
            Start-Sleep -Seconds 5
            try { 
                # Try to revive session
                $namespace.Logon("", "", $false, $false) 
            } catch {}
        }
    }

    if (-not $sourceFolder) {
        Write-Error "CRITICO: No se pudo acceder a una carpeta principal tras varios intentos. Se omitirá."
        continue
    }

    if ($excludedFolderNames -contains $sourceFolder.Name) {
        # Write-Host "  [SKIP] Saltando carpeta excluida: $($sourceFolder.Name)" -ForegroundColor DarkGray
        continue
    }

    # Ensure root folder exists in PST (Inbox, Sent Items, etc.) - ONLY IF NEEDED
    $folderRetryCount = 0
    $folderMaxRetries = 3
    $folderSuccess = $false

    while (-not $folderSuccess -and $folderRetryCount -lt $folderMaxRetries) {
        try {
            # Si estamos en un reintento, refrescar el objeto sourceFolder por si la conexión murió
            if ($folderRetryCount -gt 0) {
                 Write-Host "  [REINTENTO] Refrescando objeto de carpeta..." -ForegroundColor DarkGray
                 try { $sourceFolder = $namespace.GetFolderFromID($fid) } catch {}
            }

            if (-not $sourceFolder) { throw "No se pudo recuperar la carpeta tras desconexión." }

            # 2. OPTIMIZACIÓN CARPETAS VACIAS: Verificar si hay items antes de crear carpeta
            $itemsInBranch = Get-RecursiveItemCount -folder $sourceFolder -filter $filter -monthsToInclude $BackupMonthsList
            
            if ($itemsInBranch -gt 0) {
                $targetFolder = $pstRootFolder.Folders | Where-Object {$_.Name -eq $sourceFolder.Name}
                if (-not $targetFolder) {
                    $targetFolder = $pstRootFolder.Folders.Add($sourceFolder.Name)
                }
                
                Process-RecursiveItems -sourceFolder $sourceFolder -targetPstFolder $targetFolder -filter $filter -action $BackupAction -monthsToInclude $BackupMonthsList
            } else {
                # Write-Host "  Omitiendo carpeta vacía (sin correos en el rango): $($sourceFolder.Name)" -ForegroundColor DarkGray
            }
            $folderSuccess = $true
        } catch {
            $folderRetryCount++
            $exMsg = $_.Exception.Message
            $hresult = "0x{0:X}" -f $_.Exception.HResult
            
            # Detectar errores de conexión (0xCD840115, 0x80040115, etc)
            if ($hresult -match "CD840115" -or $hresult -match "80040115" -or $exMsg -match "red" -or $exMsg -match "network") {
                Write-Warning "Problema de conexión detectado ($hresult) en '$($sourceFolder.Name)'. Intento $folderRetryCount/$folderMaxRetries. Esperando 10s..."
                Start-Sleep -Seconds 10
                try { $namespace.Logon("", "", $false, $false) } catch {}
            } else {
                Write-Error "Error procesando carpeta '$($sourceFolder.Name)': $exMsg"
                break # Si no es error de red, no reintentamos infinitamente
            }
        }
    }
}

Write-Progress -Activity "Backup de correos de Outlook $BackupYear completado" `
               -Status "Finalizado. Total de correos procesados: $script:processedItems" `
               -PercentComplete 100 `
               -Completed

Write-Host "Backup completado. Total de correos procesados: $script:processedItems"

# Disconnect the PST file from Outlook
if ($pstStore) {
    try {
        Write-Host "Desconectando archivo PST '$backupFileName' de Outlook..."
        # La referencia $pstStore puede ser un wrapper COM que a veces falla al pasarse directamente.
        # Es mas seguro buscar el objeto Store fresco en la coleccion Stores por FilePath antes de removerlo.
        
        $storeToRemove = $null
        foreach ($s in $namespace.Stores) {
            try {
                if ($s.FilePath -eq $pstFilePath) {
                    $storeToRemove = $s
                    break
                }
            } catch {}
        }

        if ($storeToRemove) {
             $namespace.RemoveStore($storeToRemove.GetRootFolder())
             Write-Host "Archivo PST desconectado exitosamente."
        } else {
             # Si no lo encontramos por path, intentamos con el objeto original como fallback
             try {
                 $namespace.RemoveStore($pstStore.GetRootFolder())
                 Write-Host "Archivo PST desconectado exitosamente (via objeto original)."
             } catch {
                 Write-Warning "No se pudo desconectar el PST usando el objeto original. Puede que ya esté desconectado."
             }
        }

    } catch {
        Write-Host "Error al desconectar PST: $($_.Exception.Message)"
    }
}

# --- Headless Finalization ---
if ($Headless -and $LogFile) {
    $finalMeta = @{
        taskName = "OutlookBackup"
        timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        status = "Success"
        message = "Backup completed successfully. Processed $script:processedItems items."
        details = @{
            year = $BackupYear
            totalItems = $script:processedItems
            targetPath = $pstFilePath
        }
    }
    $finalMeta | ConvertTo-Json -Compress | Out-File -FilePath $LogFile -Append -Encoding UTF8 -Force
}
