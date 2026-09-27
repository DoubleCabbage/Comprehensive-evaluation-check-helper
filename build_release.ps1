param([switch]$OneFile)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    $python = Join-Path $projectRoot '.conda-env\python.exe'
}
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    throw '找不到项目 Python 环境。请先按照 README 安装 .venv。'
}

& $python -c 'import PyInstaller'
if ($LASTEXITCODE -ne 0) {
    throw '尚未安装 PyInstaller。请运行：python -m pip install pyinstaller'
}

$bundleMode = if ($OneFile) { '--onefile' } else { '--onedir' }
& $python -m PyInstaller --noconfirm --clean --windowed $bundleMode `
    --name CollegeReview --paths src `
    --add-data 'excel_navigation.ps1:.' `
    batch_review_desktop.py
if ($LASTEXITCODE -ne 0) {
    throw "打包失败，退出码：$LASTEXITCODE"
}

Write-Host "打包完成：$(Join-Path $projectRoot 'dist')" -ForegroundColor Green
