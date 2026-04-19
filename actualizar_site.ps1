# Sincronizador Maestro "Showroom Todo-en-Uno" (Version Ultra-Robusta v2.2)
# Repositorio: site_estrategia (pablorouco-ux)

# 0. Preparacion y Seguridad
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Forzar salida limpia ante errores fatales
$ErrorActionPreference = "Stop"

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   Sincronizador Showroom - Natura & Avon (Ultra-Robust)" -ForegroundColor Cyan
Write-Host "==========================================================`n" -ForegroundColor Cyan

# 1. Configuracion de Rutas
$rutaExcel = "G:\Unidades compartidas\Gestion Comercial\1. Estrategias\2026\Padres\Site Showroom\Brief Site Estrategia.xlsx"
$rutaImagenesOrigen = "G:\Unidades compartidas\Gestion Comercial\1. Estrategias\2026\Padres\Site Showroom\Imagenes"
$rutaImagenesDestino = Join-Path $PSScriptRoot "imagenes"
$archivoData = Join-Path $PSScriptRoot "data.json"

if (!(Test-Path $rutaExcel)) {
    Write-Host "[ERROR] El archivo Excel no existe en: $rutaExcel" -ForegroundColor Red
    pause; exit
}

# 2. Inicializacion de Excel COM
Write-Host "[1/3] Extrayendo datos del Excel local..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "  > Abriendo libro..." -ForegroundColor Gray
    $wb = $excel.Workbooks.Open($rutaExcel, [Type]::Missing, $true) # Abrir solo lectura
    
    # --- PROCESAMIENTO: Hoja 'Site' ---
    Write-Host "  > Procesando Hoja: Site..." -ForegroundColor Gray
    $wsSite = $wb.Sheets.Item("Site")
    $dataSite = $wsSite.UsedRange.Value2
    $siteJson = @()

    if ($null -ne $dataSite) {
        $maxRows = $dataSite.GetUpperBound(0)
        $maxCols = $dataSite.GetUpperBound(1)
        Write-Host "    (Filas: $maxRows, Columnas: $maxCols)" -ForegroundColor DarkGray

        for ($i = 2; $i -le $maxRows; $i++) {
            $marca = $dataSite[$i, 1]
            if (!$marca) { continue }
            
            $plans = @()
            for ($j = 0; $j -lt 4; $j++) {
                $colCod = 2 + ($j * 2)
                $colDesc = 3 + ($j * 2)
                
                if ($colDesc -le $maxCols) {
                    $cod = $dataSite[$i, $colCod]
                    $desc = $dataSite[$i, $colDesc]
                    if ($cod -and $desc) {
                        $plans += @{ cod = $cod.ToString().Trim(); desc = $desc.ToString().Trim(); type = [char](65 + $j) }
                    }
                }
            }
            $siteJson += @{ marca = $marca.ToString().Trim(); plans = $plans }
        }
    }

    # --- PROCESAMIENTO: Hoja 'Gerencias' ---
    Write-Host "  > Procesando Hoja: Gerencias..." -ForegroundColor Gray
    $wsGer = $wb.Sheets.Item("Gerencias")
    $dataGer = $wsGer.UsedRange.Value2
    $gerenciasJson = @()

    if ($null -ne $dataGer) {
        $maxRows = $dataGer.GetUpperBound(0)
        $maxCols = $dataGer.GetUpperBound(1)
        Write-Host "    (Filas: $maxRows, Columnas: $maxCols)" -ForegroundColor DarkGray

        for ($i = 2; $i -le $maxRows; $i++) {
            if ($maxCols -ge 2) {
                $id = $dataGer[$i, 1]
                $name = $dataGer[$i, 2]
                if ($id -and $name) {
                    $gerenciasJson += @{ id = $id.ToString().Trim(); name = $name.ToString().Trim() }
                }
            }
        }
    }

    # --- PROCESAMIENTO: Hoja 'Cuotas' (¡Con S!) ---
    Write-Host "  > Procesando Hoja: Cuotas..." -ForegroundColor Gray
    $wsCuo = $null
    try { $wsCuo = $wb.Sheets.Item("Cuotas") } catch { throw "No se encontro la hoja 'Cuotas'. Revisa el nombre." }
    
    $dataCuo = $wsCuo.UsedRange.Value2
    $cuotasJson = @{}

    if ($null -ne $dataCuo) {
        $maxRows = $dataCuo.GetUpperBound(0)
        $maxCols = $dataCuo.GetUpperBound(1)
        Write-Host "    (Filas: $maxRows, Columnas: $maxCols)" -ForegroundColor DarkGray

        for ($i = 2; $i -le $maxRows; $i++) {
            if ($maxCols -ge 17) {
                $cod = $dataCuo[$i, 2]   # Columna B
                $qty = $dataCuo[$i, 3]   # Columna C
                $zid = $dataCuo[$i, 17]  # Columna Q (17)
                
                if ($cod -and $zid) {
                    $szid = $zid.ToString().Trim()
                    $scod = $cod.ToString().Trim()
                    if (!$cuotasJson.ContainsKey($szid)) { $cuotasJson[$szid] = @{} }
                    $cuotasJson[$szid][$scod] = [int]$qty
                }
            }
        }
    }

    $finalData = @{
        updatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        site = $siteJson
        gerencias = $gerenciasJson
        cuotas = $cuotasJson
    }

    $finalData | ConvertTo-Json -Depth 10 | Out-File -FilePath $archivoData -Encoding utf8
    Write-Host "  [OK] data.json generado correctamente." -ForegroundColor Green

} catch {
    Write-Host "[ERROR] Crucial durante el procesamiento de Excel:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    pause; exit
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# 3. Sincronizacion de Imagenes (Robocopy)
if (Test-Path $rutaImagenesOrigen) {
    Write-Host "[2/3] Sincronizando imagenes..." -ForegroundColor Yellow
    robocopy $rutaImagenesOrigen $rutaImagenesDestino /MIR /NDL /NJH /NJS /nc /ns /np /R:1 /W:1
    Write-Host "  [OK] Imagenes sincronizadas." -ForegroundColor Green
}

# 4. Flujo de Git
Write-Host "[3/3] Desplegando a GitHub..." -ForegroundColor Cyan

try {
    git add .
    $status = git status --porcelain
    if ($status) {
        git commit -m "Showroom Auto-Update $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    }
    
    Write-Host "  > Sincronizando con GitHub (pull --rebase)..." -ForegroundColor Gray
    git pull --rebase origin main
    
    Write-Host "  > Subiendo cambios..." -ForegroundColor Gray
    git push origin main
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`n¡SITIO ACTUALIZADO EXITOSAMENTE!" -ForegroundColor Green
    }
} catch {
    Write-Host "`n[ERROR] Fallo en la sincronizacion con GitHub." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    git rebase --abort 2>$null
}

Write-Host "`nProceso finalizado. Presione cualquier tecla..." -ForegroundColor Gray
pause
