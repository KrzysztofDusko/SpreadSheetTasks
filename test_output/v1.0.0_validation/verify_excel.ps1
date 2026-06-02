$ErrorActionPreference = "Stop"
$dir = Join-Path $PSScriptRoot "."
$files = @(Get-ChildItem -LiteralPath $dir -Filter "*.xlsx") + @(Get-ChildItem -LiteralPath $dir -Filter "*.xlsb") | Sort-Object Name

Write-Host "Opening $($files.Count) files with Excel COM..."
Write-Host ""

$excel = $null
$totalOK = 0
$totalFail = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    foreach ($file in $files) {
        $wb = $null
        try {
            $path = $file.FullName
            $wb = $excel.Workbooks.Open($path, $null, $true)  # ReadOnly
            $sheetCount = $wb.Sheets.Count
            $name = $file.Name
            Write-Host "  [OK] $($name.PadRight(38)) sheets=$sheetCount" -ForegroundColor Green
            $totalOK++
        }
        catch {
            Write-Host "  [FAIL] $($file.Name.PadRight(38)) $($_.Exception.Message)" -ForegroundColor Red
            $totalFail++
        }
        finally {
            if ($wb) {
                $wb.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
            }
        }
    }
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host ("=" * 60)
Write-Host "OK: $totalOK  |  FAIL: $totalFail  |  Total: $($totalOK + $totalFail)"
Write-Host ("=" * 60)

if ($totalFail -gt 0) { exit 1 }
