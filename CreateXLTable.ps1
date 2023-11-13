###################################################################
# https://devblogs.microsoft.com/scripting/clean-up-your-powershell-environment-by-tracking-variable-use/

$startupVariables = ""

New-Variable -Force -Name startupVariables -Value ( Get-Variable | % {$_.Name} )

function Cleanup-Variables {

             Get-Variable | 
             Where-Object { $startupVariables -notcontains $_.Name } |
             % { Write-Host "Deleting Variable >> $($_.Name)" 
             Remove-Variable -Name "$($_.Name)" -Force -Scope "global" }

             }

###################################################################
# https://stackoverflow.com/questions/19029850/powershell-release-com-object

function Release-Ref ($ref) {

[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

}

###################################################################
#Start Logging
$LogPath = "C:\PS Scripts\Logs"
Start-Transcript -Path (Join-Path -Path $LogPath -ChildPath (($MyInvocation.MyCommand.Name).Replace("ps1","log")))

###################################################################
# https://stackoverflow.com/questions/64970809/powershell-how-to-add-table-to-excel-spreadsheet

# path to excel file parameters
$openPath = $args[0]
$savePath = $args[1]
$row = $args[2]
$col = $args[3]

#$openPath = "H:\PowerShell Scripts\Example_Report.xlsx"
#$savePath = "H:\PowerShell Scripts\Example_Report2.xlsx"
#$row = 369
#$col = "U"

$range ="A1:$col$row"

Function Create-ExcelTable($preTable, $postTable, $rowRange) {

    $Excel = New-Object -ComObject excel.application
    $Workbook = $Excel.Workbooks.Open($preTable)
    $Worksheet = $Workbook.worksheets.Item(1)

    $Table = $Worksheet.ListObjects.Add(
	       [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, # add a range
	       $Worksheet.Range($rowRange),	# set the region
	       $null,
	       [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes 	# yes, to headers
	       )

    # Save and close the workbok as an Excel file
    $Excel.DisplayAlerts = $false # Ignore / hide alerts
    $Worksheet.SaveAs($postTable,51,$null,$null,$null,$null,$null,$null,$null,'True')
    $Excel.Quit()

    #execute com object release function
    Release-Ref($worksheet)
    Release-Ref($Workbook)
    Release-Ref($Excel)

}


Create-ExcelTable $openPath $savePath $range


# execute function
Cleanup-Variables


#Stop Logging
Stop-Transcript

exit