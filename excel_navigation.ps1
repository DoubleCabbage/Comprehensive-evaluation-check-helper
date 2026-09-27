param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][string]$CellAddress
)

$ErrorActionPreference = 'Stop'

try {
    $targetPath = [System.IO.Path]::GetFullPath($WorkbookPath)
    if (-not [System.IO.File]::Exists($targetPath)) {
        throw "Workbook not found: $targetPath"
    }

    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        $excel = New-Object -ComObject Excel.Application
    }

    $excel.Visible = $true
    $workbook = $null
    foreach ($opened in $excel.Workbooks) {
        if ([string]::Equals([string]$opened.FullName, $targetPath,
                [System.StringComparison]::OrdinalIgnoreCase)) {
            $workbook = $opened
            break
        }
    }
    if ($null -eq $workbook) {
        $workbook = $excel.Workbooks.Open($targetPath, 0)
    }

    $worksheet = $workbook.Worksheets.Item($SheetName)
    $workbook.Activate()
    $worksheet.Activate()
    $target = $worksheet.Range($CellAddress)
    $excel.Goto($target, $true)
    $target.Select()

    try {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelWindowFocus {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr handle, int command);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr handle);
}
'@
        $handle = [System.IntPtr]::new([long]$excel.Hwnd)
        [ExcelWindowFocus]::ShowWindow($handle, 9) | Out-Null
        [ExcelWindowFocus]::SetForegroundWindow($handle) | Out-Null
    } catch {
        # Cell navigation succeeded even if Windows declines to change foreground focus.
    }
    exit 0
} catch {
    [System.Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
