# ==========================================
# Multiple-IP-Ping-Excel.ps1 (Safe Version)
# ==========================================

$logDir = "C:\Users\Administrator\Desktop\Ayush_Excel_Sheet\Script's\Continous\Logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
Start-Transcript -Path (Join-Path $logDir ("PingLog_{0:yyyy-MM-dd_HH-mm-ss}.txt" -f (Get-Date))) -Append

# Load Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Path of source Excel file (contains only IPs)
$sourceFile = "C:\Users\Administrator\Desktop\Ayush_Excel_Sheet\Script's\Continous\All-IP-Source-Sheet.xlsx"

# Path of result Excel file (create copy first)
$newFilePath = "C:\Users\Administrator\Desktop\Ayush_Excel_Sheet\Script's\Continous\Ping-Result_{0:yyyy-MM-dd_HH-mm-ss}.xlsx" -f (Get-Date)
Copy-Item $sourceFile $newFilePath -Force

# Open the copied file, so the source stays untouched
$workbook = $excel.Workbooks.Open($newFilePath)
$sheet = $workbook.Sheets.Item(1)

# Find last row with data
$lastRow = $sheet.Cells($sheet.Rows.Count, 1).End(-4162).Row   # -4162 = xlUp

# Add header for results
$sheet.Cells(1, 2).Value2 = "Status"

# Loop through each IP in column A
for ($row = 2; $row -le $lastRow; $row++) {
    $ip = $sheet.Cells($row, 1).Value2

    if (![string]::IsNullOrWhiteSpace($ip)) {
        try {
            $successCount = 0
            for ($i=1; $i -le 10; $i++) {
                if (Test-Connection -ComputerName $ip -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                    $successCount++
                }
                Start-Sleep -Milliseconds 100
            }
            $failCount = 10 - $successCount

            if ($successCount -gt 0) {
                Write-Host "$ip is Working ($successCount/10 replies)" -ForegroundColor Green
                $sheet.Cells($row, 2).Value2 = "Up ($successCount/10 replies)"
                $sheet.Cells($row, 2).Interior.ColorIndex = 4  # Green
            }
            else {
                Write-Host "$ip is Not Working (0/10 replies)" -ForegroundColor Red
                $sheet.Cells($row, 2).Value2 = "Down (0/10 replies)"
                $sheet.Cells($row, 2).Interior.ColorIndex = 3  # Red
            }
        }
        catch {
            Write-Host "$ip is Not Reachable (0/10 replies)" -ForegroundColor Red
            $sheet.Cells($row, 2).Value2 = "Down (0/10 replies)"
            $sheet.Cells($row, 2).Interior.ColorIndex = 3  # Red
        }

        [System.GC]::Collect()
        Start-Sleep -Milliseconds 200
    }
}

# Save changes only in the copied file
$workbook.Save()
$workbook.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "✅ Results saved to $newFilePath"
Stop-Transcript
Start-Sleep -Seconds 10
