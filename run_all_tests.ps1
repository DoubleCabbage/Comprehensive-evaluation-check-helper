param(
    [string]$ExcelFile = "",
    [string]$CasesFile = "",
    [string]$PythonExecutable = "",
    [switch]$SkipUnitTests
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectRoot

function Invoke-Python {
    param(
        [string]$StepName,
        [string[]]$CommandArgs
    )
    & $script:PythonPath @CommandArgs
    if ($LASTEXITCODE -ne 0) {
        throw "$StepName failed with exit code $LASTEXITCODE. Review the output immediately above this message."
    }
}

function Resolve-ProjectPath {
    param([string]$PathValue)
    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot $PathValue))
}

function Test-PythonModule {
    param(
        [string]$Executable,
        [string]$ModuleName
    )
    & $Executable -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ModuleName') else 1)" *> $null
    return ($LASTEXITCODE -eq 0)
}

Write-Host "Activity score matcher: one-click test" -ForegroundColor Cyan
Write-Host "Project: $ProjectRoot"

if ($PythonExecutable) {
    $script:PythonPath = Resolve-ProjectPath $PythonExecutable
    if (-not (Test-Path -LiteralPath $script:PythonPath -PathType Leaf)) {
        throw "Python executable not found: $script:PythonPath"
    }
    if (-not (Test-PythonModule -Executable $script:PythonPath -ModuleName "openpyxl")) {
        throw "The selected Python does not have openpyxl: $script:PythonPath"
    }
    Write-Host "Using Python: $script:PythonPath"
} else {
    $EnvironmentPath = Join-Path $ProjectRoot ".conda-env"
    $ProjectPython = Join-Path $EnvironmentPath "python.exe"
    $BundledPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    $script:PythonPath = ""

    if (
        (Test-Path -LiteralPath $ProjectPython -PathType Leaf) -and
        (Test-PythonModule -Executable $ProjectPython -ModuleName "openpyxl")
    ) {
        $script:PythonPath = $ProjectPython
        Write-Host "Using project Conda environment: $EnvironmentPath"
    } elseif (
        (Test-Path -LiteralPath $BundledPython -PathType Leaf) -and
        (Test-PythonModule -Executable $BundledPython -ModuleName "openpyxl")
    ) {
        $script:PythonPath = $BundledPython
        Write-Host "Project Conda dependencies are incomplete; using the offline bundled Python." -ForegroundColor Yellow
        Write-Host "Using Python: $script:PythonPath"
    } else {
        $CondaCommand = Get-Command conda -ErrorAction SilentlyContinue
        if (-not $CondaCommand) {
            throw "Conda was not found. Install Conda or pass -PythonExecutable."
        }
        if (-not (Test-Path -LiteralPath $ProjectPython -PathType Leaf)) {
            Write-Host "First run: creating the project Conda environment..." -ForegroundColor Yellow
            & $CondaCommand.Source create --prefix $EnvironmentPath python=3.11 pip -y
            if ($LASTEXITCODE -ne 0) {
                throw "Conda environment creation failed. Check the network and Conda configuration."
            }
        }
        Write-Host "Installing openpyxl and pytest in the project Conda environment..." -ForegroundColor Yellow
        & $CondaCommand.Source install --prefix $EnvironmentPath openpyxl pytest -y
        if ($LASTEXITCODE -ne 0) {
            throw "Dependency installation failed. Connect to the network and run again."
        }
        $script:PythonPath = $ProjectPython
        Write-Host "Using project Conda environment: $EnvironmentPath"
    }
}

if (-not $SkipUnitTests) {
    Write-Host "`n[1/4] Running code tests" -ForegroundColor Cyan
    if (Test-PythonModule -Executable $script:PythonPath -ModuleName "pytest") {
        Invoke-Python -StepName "Code tests" -CommandArgs @("-m", "pytest")
    } else {
        Write-Host "pytest is unavailable; running the offline smoke-check suite." -ForegroundColor Yellow
        Invoke-Python -StepName "Offline smoke checks" -CommandArgs @("tests\smoke_checks.py")
    }
} else {
    Write-Host "`n[1/4] Code tests skipped"
}

$RawDirectory = Join-Path $ProjectRoot "data\raw"
if ($ExcelFile) {
    $ResolvedExcel = Resolve-ProjectPath $ExcelFile
    if (-not (Test-Path -LiteralPath $ResolvedExcel -PathType Leaf)) {
        throw "Excel file not found: $ResolvedExcel"
    }
    $ExcelFiles = @(Get-Item -LiteralPath $ResolvedExcel)
} else {
    $ExcelFiles = @(
        Get-ChildItem -LiteralPath $RawDirectory -File |
            Where-Object {
                $_.Extension.ToLowerInvariant() -in @(".xlsx", ".xlsm") -and
                -not $_.Name.StartsWith("~$")
            } |
            Sort-Object Name
    )
}

if ($ExcelFiles.Count -eq 0) {
    throw "No .xlsx or .xlsm files were found in data\raw."
}

$ResolvedCases = ""
if ($CasesFile) {
    $ResolvedCases = Resolve-ProjectPath $CasesFile
    if (-not (Test-Path -LiteralPath $ResolvedCases -PathType Leaf)) {
        throw "Retrieval test case file not found: $ResolvedCases"
    }
}

Write-Host "`n[2/4] Found $($ExcelFiles.Count) Excel file(s)" -ForegroundColor Cyan

$CompletedRuns = @()
foreach ($File in $ExcelFiles) {
    $RunName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $ProcessedDirectory = Join-Path $ProjectRoot "data\processed\runs\$RunName"
    $InterimDirectory = Join-Path $ProjectRoot "data\interim\runs\$RunName"
    New-Item -ItemType Directory -Force -Path $ProcessedDirectory, $InterimDirectory | Out-Null

    $CsvPath = Join-Path $ProcessedDirectory "historical_activities.csv"
    $TreePath = Join-Path $ProcessedDirectory "activity_tree.json"
    $ImportReport = Join-Path $InterimDirectory "import_report.json"
    $RetrievalReport = Join-Path $InterimDirectory "retrieval_report.json"

    Write-Host "`nProcessing: $($File.Name)" -ForegroundColor Green
    Write-Host "[3/4] Inspecting and cleaning data"
    Invoke-Python -StepName "Workbook inspection" -CommandArgs @("run.py", "inspect", $File.FullName)
    Invoke-Python -StepName "Workbook cleaning" -CommandArgs @(
        "run.py", "clean", $File.FullName,
        "--output", $CsvPath,
        "--report", $ImportReport,
        "--tree", $TreePath
    )

    Write-Host "[4/4] Running retrieval smoke tests"
    $ExperimentArgs = @(
        "experiments\baseline_retrieval.py",
        "--catalog", $CsvPath,
        "--output", $RetrievalReport,
        "--summary-only"
    )
    if ($ResolvedCases) {
        $ExperimentArgs += @("--cases", $ResolvedCases)
    }
    Invoke-Python -StepName "Retrieval evaluation" -CommandArgs $ExperimentArgs

    $CompletedRuns += [PSCustomObject]@{
        Excel = $File.FullName
        CleanData = $CsvPath
        ActivityTree = $TreePath
        ImportReport = $ImportReport
        RetrievalReport = $RetrievalReport
    }
}

Write-Host "`nAll tests completed." -ForegroundColor Green
$CompletedRuns | Format-List
