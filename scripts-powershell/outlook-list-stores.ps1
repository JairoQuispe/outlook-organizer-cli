param (
    [Parameter(Mandatory=$false)]
    [switch]$Json
)

function Format-StoreBytes {
    param($bytes)

    if (Get-Command Format-Bytes -ErrorAction SilentlyContinue) {
        return (Format-Bytes $bytes)
    }

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

function Get-StoreInfo {
    param(
        $store,
        [switch]$IncludeRawStore
    )

    $path = $null
    try { $path = $store.FilePath } catch {}

    $fileSize = $null
    try {
        if ($path -and (Test-Path $path)) {
            $fileInfo = Get-Item $path
            $fileSize = Format-StoreBytes $fileInfo.Length
        }
    } catch {}

    $exchangeStoreType = $null
    try { $exchangeStoreType = $store.ExchangeStoreType } catch {}

    $id = $null
    try { $id = $store.StoreID } catch {}

    $storeInfo = [pscustomobject]@{
        displayName = $store.DisplayName
        storeId = $id
        filePath = $path
        fileSize = $fileSize
        exchangeStoreType = $exchangeStoreType
    }

    if ($IncludeRawStore) {
        return [pscustomobject]@{
            info = $storeInfo
            store = $store
        }
    }

    return $storeInfo
}

function Get-ConnectedOutlookStores {
    param(
        $namespace,
        [switch]$IncludeRawStore
    )

    $result = @()
    foreach ($store in $namespace.Stores) {
        $result += Get-StoreInfo -store $store -IncludeRawStore:$IncludeRawStore
    }

    return $result
}

function Get-StoreByIdOrDefault {
    param(
        $namespace,
        [string]$storeId
    )

    if ($storeId) {
        foreach ($s in $namespace.Stores) {
            try {
                if ($s.StoreID -eq $storeId) { return $s }
            } catch {}
        }
        throw "No se encontró el StoreId solicitado."
    }

    try {
        return $namespace.DefaultStore
    } catch {
        return $null
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    try {
        $outlook = $null
        $namespace = $null
        $maxRetries = 3
        $retryDelay = 2

        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                try {
                    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
                    $namespace = $outlook.GetNamespace("MAPI")
                    try { $namespace.Logon("", "", $false, $false) } catch {}
                } catch {
                    $outlook = New-Object -ComObject Outlook.Application
                    $namespace = $outlook.GetNamespace("MAPI")
                    try { $namespace.Logon("", "", $false, $true) } catch {}
                }

                if (-not $namespace) {
                    throw "No se pudo obtener el Namespace MAPI de Outlook."
                }

                break
            } catch {
                $errorCode = $_.Exception.HResult
                if ($errorCode -eq 0xCC54011D -or $errorCode -eq -864313571) {
                    if ($attempt -lt $maxRetries) {
                        Write-Warning "Outlook ocupado (intento $attempt/$maxRetries). Reintentando en $retryDelay segundos..."
                        Start-Sleep -Seconds $retryDelay
                        continue
                    }
                }
                throw
            }
        }

        $stores = Get-ConnectedOutlookStores -namespace $namespace

        if ($Json) {
            # Forzar que stores SIEMPRE se serialice como arreglo (incluido 1 elemento).
            $storesArr = @($stores)
            if ($storesArr.Count -eq 0) {
                $storesJson = "[]"
            } elseif ($storesArr.Count -eq 1) {
                $singleJson = $storesArr[0] | ConvertTo-Json -Compress -Depth 6
                $storesJson = "[$singleJson]"
            } else {
                $storesJson = $storesArr | ConvertTo-Json -Compress -Depth 6
            }
            Write-Output ('{"type":"stores","stores":' + $storesJson + '}')
        } else {
            $i = 1
            foreach ($s in $stores) {
                $pathDisplay = if ($s.filePath) { $s.filePath } else { "Modo Online / Sin ruta local" }
                $idDisplay = if ($s.storeId) { $s.storeId } else { "Sin StoreID" }
                Write-Output "[$i] $($s.displayName) - ID: $idDisplay - ($pathDisplay)"
                $i++
            }
        }
    } catch {
        Write-Error "Error en outlook-list-stores.ps1: $($_.Exception.Message)"
        exit 1
    }
}
