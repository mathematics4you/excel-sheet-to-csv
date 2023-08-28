$xls = "$PSScriptRoot\example workbook.xlsx"
$csvpath = "$PSScriptRoot\sheets\"

If ("False" -eq (Test-Path -Path $xls -PathType leaf)) {
    echo "The source path does not exist or is not accessible"
    exit
}
If ("False" -eq (Test-Path -Path $csvpath)) {
    echo "The destination path does not exist or is not accessible"
    exit
}

$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $False
$objExcel.DisplayAlerts = $False

$WorkBook = $objExcel.Workbooks.Open($xls)
$xlCSV = 6
foreach($WorkSheet in $WorkBook.sheets) {
    $SheetName=$WorkSheet.Name
    $csv="$csvpath$SheetName.csv"
    $WorkSheet.SaveAs($csv,$xlCSV)
}
$SheetsList = $WorkBook.sheets | Select-Object -Property Name

$objExcel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
