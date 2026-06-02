param(
    [Parameter(Mandatory = $false)]
    [string]$TestDir = $PSScriptRoot,
    [string]$LogFile = ""
)

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Output "SKIP||Excel not installed or COM unavailable"
    exit 0
}

$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

$totalElapsed = [System.Diagnostics.Stopwatch]::StartNew()

# ---- Precheck: verify Excel can actually open a file ----
$knownGood = Get-ChildItem -Path $TestDir -Filter "*.xlsx" | Select-Object -First 1
if ($knownGood) {
    try {
        $checkWb = $excel.Workbooks.Open($knownGood.FullName, $false, $true)
        $checkSheets = $checkWb.Sheets.Count
        $checkWb.Close($false)
        Write-Output "PRECHECK|PASS|Excel responsive ($checkSheets sheets in test file)"
    } catch {
        $checkErr = $_.Exception.Message -replace "`r|`n", " "
        Write-Output "PRECHECK|FAIL|$checkErr"
        $excel.Quit()
        exit 1
    }
} else {
    Write-Output "PRECHECK|WARN|No reference file found, skipping precheck"
}

# ---- Validate all test files ----
$files = Get-ChildItem -Path $TestDir -Filter "*.xls?" -File | Sort-Object Name
$processed = 0
$passed = 0
$failed = 0

foreach ($file in $files) {
    $processed++
    $fileTimer = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $wb = $excel.Workbooks.Open($file.FullName, $false, $true)
        $sheetCount = $wb.Sheets.Count
        $wb.Close($false)
        $fileTimer.Stop()
        $ms = $fileTimer.ElapsedMilliseconds
        $passed++
        $msg = "PASS|$($file.Name)|$sheetCount sheets|${ms}ms"
        Write-Output $msg
        if ($LogFile) { Add-Content -Path $LogFile -Value $msg }
    } catch {
        $fileTimer.Stop()
        $ms = $fileTimer.ElapsedMilliseconds
        $err = $_.Exception.Message -replace "`r|`n", " "
        $failed++
        $msg = "FAIL|$($file.Name)|$err|${ms}ms"
        Write-Output $msg
        if ($LogFile) { Add-Content -Path $LogFile -Value $msg }
    }
}

$totalElapsed.Stop()
$totalSec = [math]::Round($totalElapsed.Elapsed.TotalSeconds, 2)
Write-Output "ELAPSED|$processed|$passed|$failed|${totalSec}s"
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
