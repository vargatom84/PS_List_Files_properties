$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false
$objExcel.DisplayAlerts = $False
$WorkBook = $objExcel.Workbooks.Open("C:\Test\test.xlsx")
$WorkSheet = $WorkBook.sheets.item("Munka1")
$WorkSheet.Cells.Item(1,1).value ="File"
$WorkSheet.Cells.Item(1,2).value ="Mappa"
$WorkSheet.Cells.Item(1,3).value ="Full"
$WorkSheet.Cells.Item(1,4).value ="Creation Date"
$WorkSheet.Cells.Item(1,5).value ="Last modification"
cls
$Path = "C:\Program Files (x86)\AppGate"
$Elemek = Get-ChildItem -File -Path $Path -Recurse
[int]$Row = 2
 foreach ($elem in $Elemek) {
   $WorkSheet.Cells.Item($Row,1).value = $elem.Name.ToString()
   $WorkSheet.Cells.Item($Row,2).value = $elem.DirectoryName.ToString()
   $WorkSheet.Cells.Item($Row,3).value = $elem.FullName.ToString()
   $WorkSheet.Cells.Item($Row,4).value = $elem.CreationTime.ToString()
   $WorkSheet.Cells.Item($Row,5).value = $elem.LastWriteTime.ToString()
   $Row++
} 
$WorkBook.save()
$WorkBook.Close()
$objExcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook)
#spps -n Excel
Remove-Variable objExcel
[System.GC]::Collect()