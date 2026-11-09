Param(
    [string]$AppName = "DoiChieuDuLieuHIS",
    [string]$EntryPoint = "main.py"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pythonHome = "C:\Users\Long An\AppData\Local\Python\pythoncore-3.14-64"
$pythonExe = Join-Path $projectRoot "env\Scripts\python.exe"
$tclDir = Join-Path $pythonHome "tcl\tcl8.6"
$tkDir = Join-Path $pythonHome "tcl\tk8.6"
$tkHook = Join-Path $projectRoot "rthook_tk.py"
$hooksDir = Join-Path $projectRoot "hooks"
$pythonDllDir = Join-Path $pythonHome "DLLs"

if (-not (Test-Path $pythonExe)) {
    throw "Khong tim thay Python trong virtual env: $pythonExe"
}
if (-not (Test-Path $pythonHome)) {
    throw "Khong tim thay Python Home: $pythonHome"
}

if (-not (Test-Path $EntryPoint)) {
    throw "Khong tim thay file entry point: $EntryPoint"
}
if (-not (Test-Path $tkHook)) {
    throw "Khong tim thay runtime hook: $tkHook"
}
if (-not (Test-Path $hooksDir)) {
    throw "Khong tim thay custom hooks: $hooksDir"
}
if (-not (Test-Path $tclDir)) {
    throw "Khong tim thay Tcl runtime: $tclDir"
}
if (-not (Test-Path $tkDir)) {
    throw "Khong tim thay Tk runtime: $tkDir"
}
if (-not (Test-Path $pythonDllDir)) {
    throw "Khong tim thay thu muc DLL cua Python: $pythonDllDir"
}

Write-Host "[1/4] Kiem tra PyInstaller..." -ForegroundColor Cyan
& $pythonExe -m PyInstaller --version | Out-Host

Write-Host "[2/4] Xoa output cu (neu co)..." -ForegroundColor Cyan
$buildDir = Join-Path $projectRoot "build"
$distDir = Join-Path $projectRoot "dist"
$specFile = Join-Path $projectRoot "$AppName.spec"

if (Test-Path $buildDir) { Remove-Item -LiteralPath $buildDir -Recurse -Force }
if (Test-Path $distDir) { Remove-Item -LiteralPath $distDir -Recurse -Force }
if (Test-Path $specFile) { Remove-Item -LiteralPath $specFile -Force }

Write-Host "[3/4] Build EXE..." -ForegroundColor Cyan
$env:TCL_LIBRARY = $tclDir
$env:TK_LIBRARY = $tkDir
$env:PATH = "$pythonDllDir;$env:PATH"
& $pythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name $AppName `
    --additional-hooks-dir $hooksDir `
    --hidden-import tkinter `
    --hidden-import _tkinter `
    --hidden-import tkinter.ttk `
    --hidden-import tkinter.filedialog `
    --hidden-import tkinter.font `
    --hidden-import tkinter.constants `
    --collect-all customtkinter `
    --collect-all tkcalendar `
    --collect-submodules tkinter `
    --collect-data tkinter `
    --runtime-hook $tkHook `
    --add-data "${tclDir};_tcl_data" `
    --add-data "${tkDir};_tk_data" `
    $EntryPoint

Write-Host "[4/4] Hoan tat." -ForegroundColor Green
$exePath = Join-Path $distDir "$AppName.exe"
if (Test-Path $exePath) {
    $exeInfo = Get-Item $exePath
    Write-Host "EXE: $($exeInfo.FullName)"
    Write-Host "Size: $([math]::Round($exeInfo.Length / 1MB, 2)) MB"
} else {
    throw "Build xong nhung khong tim thay file EXE o: $exePath"
}
