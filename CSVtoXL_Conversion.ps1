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
# function to replace commas in csv file with blank char > they cause issues when file is converted to xlsx
Function Remove-Commas($csv_file) {
	
	$content = Get-Content $csv_file
	$newContent = $content -replace ',' , ''
	$newContent | Set-Content -Path $csv_file

}

###################################################################
#Start Logging
$LogPath = "C:\PS Scripts\Logs"
Start-Transcript -Path (Join-Path -Path $LogPath -ChildPath (($MyInvocation.MyCommand.Name).Replace("ps1","log")))

###################################################################
# https://code.adonline.id.au/csv-to-xlsx-powershell/

# parameters
$csv = $args[0]
$xlsx = $args[1]

# define csv files
#$example_csv =  "H:\PowerShell Scripts\test.csv"

# define the xlsx files
#$example_xlsx =  "H:\PowerShell Scripts\Example_Report.xlsx"

#define the csv files delimiter
$delim = "|"

Function Convert($csv_file, $xlsx_file) {

    # Create a new Excel Workbook with one empty sheet
    $excel = New-Object -ComObject excel.application
    $workbook = $excel.Workbooks.Add(1)
    $worksheet = $workbook.worksheets.Item(1)

    # Build the query tables > add command and reformat the data
    $TextConnector = ("TEXT;" + $csv_file)
    $Connector = $worksheet.QueryTables.add($TextConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = $delim
    $query.TextFileParseType = 1
    $query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    $query.TextFileCommaDelimiter = 1

    # Execute & delete the import query
    $query.Refresh()
    $query.Delete()


    # Save and close the workbok as an Excel file
    $workbook.SaveAs($xlsx_file,51)
    $excel.quit()


    #execute com object release function
    Release-Ref($query)
    Release-Ref($Connector)
    Release-Ref($worksheet)
    Release-Ref($workbook)
    Release-Ref($excel)

}

#delete the previous xlsx files if they exist
if (Test-Path $xlsx) { del $xlsx }
#if (Test-Path $example_xlsx) { del $example_xlsx }


# replace commans
Remove-Commas $csv


# convert the csv file to xlsx
Write-Host "$csv is being converted to $xlsx"
Convert $csv $xlsx
#Convert $example_csv $example_xlsx


#delete csv file
Write-Host "$csv is being deleted"
Remove-Item $csv


# execute function
Cleanup-Variables


#Stop Logging
Stop-Transcript

exit