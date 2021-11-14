##variables
$date = Get-Date -f "MM/dd/yyyy HH:mm"
$sheet = (Get-Culture).DateTimeFormat.GetMonthName(8)
#$approval = tbd
$name = [Environment]::UserName
$hashes = Get-FileHash -Algorithm SHA256 -Path (Get-ChildItem "C:\Users\$name\Desktop\File Transfer\*.*" -Recurse -File -Force -ea SilentlyContinue -ev errs)

##excel file COM
$ExcelPath = "C:\Users\zacha\Desktop\TestFolder\File Transfer Log.xlsx" #R:\
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
$ExcelWorkBook = $Excel.Workbooks.Open($ExcelPath)
$ExcelWorkSheet = $Excel.WorkSheets.item($sheet)
$ExcelWorkSheet.activate()

##push to excel
for (($i = 0); $i -lt $hashes.Hash.Length; $i++)
{
    $nextRow = $ExcelWorkSheet.UsedRange.rows.count + 1
    $ExcelWorkSheet.Cells.Item($nextRow,1) = $date
    $ExcelWorkSheet.Cells.Item($nextRow,2) = #$approval
    $ExcelWorkSheet.Cells.Item($nextRow,3) = $name
    $ExcelWorkSheet.Cells.Item($nextRow,4) = $hashes.Path[$i]
    $ExcelWorkSheet.Cells.Item($nextRow,5) = $hashes.Hash[$i]
}

##save, close, release
$ExcelWorkSheet.UsedRange.Columns.Autofit()
$ExcelWorkBook.Save()
$ExcelWorkBook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Stop-Process -Name EXCEL -Force
