$ErrorActionPreference = 'Stop'

Function Convert-CsvInBatch
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)][String]$Folder
	)
	$ExcelFiles = Get-ChildItem -Path $Folder -Include *.xlsx, *.xlsm, *.xls -Recurse

	$excelApp = New-Object -ComObject Excel.Application
	$excelApp.DisplayAlerts = $false

	$ExcelFiles | ForEach-Object {
		$workbook = $excelApp.Workbooks.Open($_.FullName)
		$csvFilePath = $_.FullName -Replace "\.xlsx", ".csv" ` -Replace "\.xls", ".csv" ` -Replace "\.xlsm", ".csvm"
		$workbook.SaveAs($csvFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
		$workbook.Close($false)
	}
	$excelApp.Workbooks.Close()
	$excelApp.Visible = $true
	Start-Sleep 5
	$excelApp.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
}
$FolderPath = "C:\Users\aarel\Downloads\xls-csv"

Convert-CsvInBatch -Folder $FolderPath