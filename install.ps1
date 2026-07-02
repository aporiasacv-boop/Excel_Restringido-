param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Src = Join-Path $ProjectRoot "src"
$OutDir = Join-Path $ProjectRoot "LISTOS"
$TempDir = Join-Path $env:LOCALAPPDATA "Temp\OlnaturaInstall"
if (Test-Path $TempDir) { Remove-Item -LiteralPath $TempDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$StdModules = @(
    "modAppConstants.bas",
    "modUtils.bas",
    "modApi.bas",
    "modBootstrap.bas",
    "modLogin.bas",
    "modAdmin.bas"
)

$ThisWorkbookCode = @"
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo OpenFailed
    InitializeApplication
    modLogin.ShowLoginForm
    Exit Sub
OpenFailed:
    MsgBox "Error al iniciar: " & Err.Description, vbCritical, "Olnatura"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modApi.ClearSession
    modAdmin.HideUsersSheet
    modAdmin.HideAdminButtons
End Sub
"@

function Prepare-VbaFile {
    param([string]$FileName)
    $source = Join-Path $Src $FileName
    if (-not (Test-Path -LiteralPath $source)) { Write-Error "Falta: $source" }
    $content = [System.IO.File]::ReadAllText($source, [System.Text.UTF8Encoding]::new($false))
    $content = $content -replace "`r?`n", "`r`n"
    $dest = Join-Path $TempDir $FileName
    [System.IO.File]::WriteAllText($dest, $content, [System.Text.Encoding]::GetEncoding(1252))
    return $dest
}

function Get-ModuleCount($vbProj) {
    $n = 0
    foreach ($c in @($vbProj.VBComponents)) {
        if ($c.Type -eq 1) { $n++ }
    }
    return $n
}

function Assert-ComponentType {
    param($Component, [int]$ExpectedType, [string]$Label)
    if ($Component.Type -ne $ExpectedType) {
        Write-Error "$Label importado como tipo $($Component.Type), esperado $ExpectedType"
    }
}

function Clear-ReadOnlyFlag {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.IsReadOnly) { $item.IsReadOnly = $false }
}

function Remove-OlnaturaAdminButtons {
    param($Worksheet)
    foreach ($shape in @($Worksheet.Shapes)) {
        if ($shape.Name -like "Olnatura_Btn_*") {
            $shape.Delete()
        }
    }
}

function Add-OlnaturaAdminButtons {
    param($Worksheet)
    Remove-OlnaturaAdminButtons $Worksheet
    $xlButtonControl = 0
    $msoFalse = 0
    $left = 8
    $width = 118
    $height = 22
    $defs = @(
        @{ Name = "Olnatura_Btn_Panel"; Macro = "AdministrarUsuarios"; Label = "Colaboradores"; Top = 8 },
        @{ Name = "Olnatura_Btn_Alta"; Macro = "AltaUsuario"; Label = "Alta usuario"; Top = 34 },
        @{ Name = "Olnatura_Btn_Baja"; Macro = "BajaUsuario"; Label = "Baja usuario"; Top = 60 },
        @{ Name = "Olnatura_Btn_Reactivar"; Macro = "ReactivarUsuario"; Label = "Reactivar"; Top = 86 }
    )
    foreach ($def in $defs) {
        $shape = $Worksheet.Shapes.AddFormControl($xlButtonControl, $left, $def.Top, $width, $height)
        $shape.Name = $def.Name
        $shape.OnAction = $def.Macro
        $shape.TextFrame.Characters().Text = $def.Label
        $shape.Visible = $msoFalse
    }
}

$excelProcs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excelProcs -and -not $Force) {
    Write-Host ""
    Write-Host "Excel esta abierto. Cierre Excel y vuelva a ejecutar 2_INSTALAR_EXCEL.bat" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$InputDir = Join-Path $ProjectRoot "entrada"
New-Item -ItemType Directory -Path $InputDir -Force | Out-Null

$Workbooks = @()
$Workbooks += Get-ChildItem -LiteralPath $InputDir -Filter "NIKZON*.xlsm" -File -ErrorAction SilentlyContinue
$Workbooks += Get-ChildItem -LiteralPath $ProjectRoot -Filter "NIKZON*.xlsm" -File -ErrorAction SilentlyContinue
$Workbooks = @($Workbooks | Sort-Object FullName -Unique)

if ($Workbooks.Count -eq 0) {
    Write-Host ""
    Write-Host "ERROR: No hay archivos NIKZON*.xlsm para instalar." -ForegroundColor Red
    Write-Host ""
    Write-Host "Copie los Excel originales a:" -ForegroundColor Yellow
    Write-Host "  $InputDir"
    Write-Host ""
    Write-Host "Ejemplo: NIKZON 1.xlsm, NIKZON 2.xlsm (los 4 formatos)"
    Write-Host "Luego ejecute de nuevo 2_INSTALAR_EXCEL.bat"
    Write-Host ""
    exit 1
}

Write-Host "Origen: $($Workbooks.Count) archivo(s) NIKZON"

if ($Force) {
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
try { $excel.Calculation = -4135 } catch { }  # xlCalculationManual
$xlOpenXMLWorkbookMacroEnabled = 52

try {
    foreach ($wbFile in $Workbooks) {
        Write-Host ""
        Write-Host "Instalando: $($wbFile.Name)"

        $sourcePath = $wbFile.FullName
        Clear-ReadOnlyFlag $sourcePath

        $workCopy = Join-Path $TempDir $wbFile.Name
        Write-Host "  Copiando a temp..."
        Copy-Item -LiteralPath $sourcePath -Destination $workCopy -Force
        Clear-ReadOnlyFlag $workCopy
        if (-not (Test-Path -LiteralPath $workCopy)) {
            Write-Error "No se pudo copiar a temp: $sourcePath"
        }

        Write-Host "  Abriendo en Excel (archivos grandes pueden tardar 1-2 min)..."
        $book = $excel.Workbooks.Open($workCopy, 0, $false)
        Write-Host "  Abierto."
        if ($null -eq $book.VBProject) {
            Write-Error "Sin acceso VBA. Activa 'Confiar en el acceso al modelo de objetos de VBA' en Excel."
        }
        $vbProj = $book.VBProject

        Write-Host "  Limpiando modulos viejos..."
        $removeList = @()
        foreach ($comp in @($vbProj.VBComponents)) {
            if ($comp.Type -eq 1 -or $comp.Type -eq 2 -or $comp.Type -eq 3) {
                $removeList += $comp.Name
            }
        }
        foreach ($name in $removeList) {
            $vbProj.VBComponents.Remove($vbProj.VBComponents.Item($name))
        }

        foreach ($fileName in $StdModules) {
            $path = Prepare-VbaFile $fileName
            $comp = $vbProj.VBComponents.Import($path)
            Assert-ComponentType $comp 1 $fileName
            Write-Host "  + $($comp.Name)"
        }

        Write-Host "  Guardando en LISTOS\..."

        if ((Get-ModuleCount $vbProj) -ne $StdModules.Count) {
            Write-Error "Modulos incompletos en $($wbFile.Name)"
        }

        $codeMod = $vbProj.VBComponents.Item("ThisWorkbook").CodeModule
        if ($codeMod.CountOfLines -gt 0) {
            $codeMod.DeleteLines(1, $codeMod.CountOfLines)
        }
        $codeMod.AddFromString($ThisWorkbookCode)

        Write-Host "  Botones Admin (esquina superior izquierda)..."
        Add-OlnaturaAdminButtons $book.Worksheets.Item(1)

        $outPath = Join-Path $OutDir $wbFile.Name
        if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
        $book.SaveAs($outPath, $xlOpenXMLWorkbookMacroEnabled)
        $book.Close($false)
        Clear-ReadOnlyFlag $outPath

        try {
            Copy-Item -LiteralPath $outPath -Destination $sourcePath -Force
            Clear-ReadOnlyFlag $sourcePath
            Write-Host "  OK"
        }
        catch {
            Write-Host "  OK -> LISTOS\$($wbFile.Name)" -ForegroundColor Green
        }
    }
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if (Test-Path $TempDir) { Remove-Item -LiteralPath $TempDir -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
Write-Host "Listo: $($Workbooks.Count) archivo(s) en LISTOS\"
Write-Host "Suba esos .xlsm a OneDrive para produccion."
Write-Host "Login: Admin / Admin123!"
