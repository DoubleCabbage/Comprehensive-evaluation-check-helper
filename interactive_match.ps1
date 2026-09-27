param(
    [string]$Catalog = "",
    [int]$Top = 5
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectRoot

$ProjectPython = Join-Path $ProjectRoot ".conda-env\python.exe"
$BundledPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if (Test-Path -LiteralPath $ProjectPython -PathType Leaf) {
    $PythonPath = $ProjectPython
} elseif (Test-Path -LiteralPath $BundledPython -PathType Leaf) {
    $PythonPath = $BundledPython
} else {
    $PythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $PythonCommand) {
        throw "Python was not found. Run run_all_tests.cmd first."
    }
    $PythonPath = $PythonCommand.Source
}

$CommandArgs = @("interactive_match.py", "--top", $Top)
if ($Catalog) {
    $CommandArgs += @("--catalog", $Catalog)
}

& $PythonPath @CommandArgs
exit $LASTEXITCODE
