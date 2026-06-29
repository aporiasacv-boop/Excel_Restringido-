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
    "modLogin.bas"
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

$excelProcs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excelProcs -and -not $Force) {
    Write-Host ""
    Write-Host "Excel esta abierto. Cierra Excel o ejecuta INSTALAR.bat" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$Workbooks = Get-ChildItem -LiteralPath $ProjectRoot -Filter "NIKZON*.xlsm" -File
if ($Workbooks.Count -eq 0) { Write-Error "No hay archivos NIKZON*.xlsm en $ProjectRoot" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$xlOpenXMLWorkbookMacroEnabled = 52

try {
    foreach ($wbFile in $Workbooks) {
        Write-Host "Instalando: $($wbFile.Name)"

        $sourcePath = $wbFile.FullName
        Clear-ReadOnlyFlag $sourcePath

        $workCopy = Join-Path $TempDir $wbFile.Name
        Copy-Item -LiteralPath $sourcePath -Destination $workCopy -Force
        Clear-ReadOnlyFlag $workCopy

        $book = $excel.Workbooks.Open($workCopy, $false, $false)
        if ($null -eq $book.VBProject) {
            Write-Error "Sin acceso VBA. Activa 'Confiar en el acceso al modelo de objetos de VBA' en Excel."
        }
        $vbProj = $book.VBProject

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

        if ((Get-ModuleCount $vbProj) -ne $StdModules.Count) {
            Write-Error "Modulos incompletos en $($wbFile.Name)"
        }

        $codeMod = $vbProj.VBComponents.Item("ThisWorkbook").CodeModule
        if ($codeMod.CountOfLines -gt 0) {
            $codeMod.DeleteLines(1, $codeMod.CountOfLines)
        }
        $codeMod.AddFromString($ThisWorkbookCode)

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
Write-Host "Excel listo en LISTOS. Configure API_BASE_URL en modAppConstants si despliega la API."
Write-Host "Login central: Admin / Admin123!"
