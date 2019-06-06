#*******************************************************************************
#* ORSEXCEL.ps1
#* Read delimited file and convert it to Excel xlsx table format.
#* Carlos Kassab
#* 2019-June-06
#******************************************************************************* 


# Install Module ImportExcel, nuget package, it is needed to authorize Powershell to install nuget.
#
# Go to https://www.powershellgallery.com/packages/ImportExcel/6.0.0
#
# Open PowerShell console as Administrator.
#
# Search and install ImportExcel module:
# Find-Module ImportExcel | Install-Module
#
# Get-Module ImportExcel #to see all module details
# Get-Command -Module ImportExcel #to see all module commands
#
# Set execution policy in order to use powershell scripts:
# Open powershell console as adminitrator, then run the next command:
# Set-ExecutionPolicy -ExecutionPolicy Unrestricted
#
# For script testing open a command window and run:
# ORSEXCEL.bat \OpenReportingSystem\SampleData\tcsqlpur001 \OpenReportingSystem\SampleData\tcsqlpur001.xlsx "Purchase Report"

if( $args.count -lt 1 ) 
{
	Write-output "Usage: ORSEXCEL.bat delimitedfile excelfile.xlsx 'report title'"
	exit
} else {
	$sourceFileName = $($args[0])
	$targetFileName = $($args[1])
	$reportTitle = $($args[2])	
}

# Delete possible old file.
Remove-Item $targetFileName -ErrorAction Ignore 

# Reading LN Delimited file
$myReportData = Get-Content $sourceFileName | Where-Object { !$_.StartsWith("#") } | ConvertFrom-Csv -Delimiter "|" 

# Create Excel object and store it in a variable.
$myExcelReport = $myReportData | Export-Excel $targetFileName -AutoSize -AutoFilter -Title $reportTitle -TitleBold -TitleSize 14 -PassThru

# Formatting WorkSheet
$workSheet1 = $myExcelReport.Workbook.WorkSheets[1]

# Get title end column, considering a maximum of 78 report columns
$reportColumns = ($myReportData | get-member -type NoteProperty).count
if( $reportColumns -le 26 ) { $titleColumnEnd = "$([char]$( 65 + ( $reportColumns - 1 ) ) )1" }
if( $reportColumns -gt 26 -And $reportColumns -le 52 ) { $titleColumnEnd = "A$([char]$( 65 + ( $reportColumns - 27 ) ) )1" }
if( $reportColumns -gt 52 -And $reportColumns -le 78 ) { $titleColumnEnd = "B$([char]$( 65 + ( $reportColumns - 53 ) ) )1" }

Set-ExcelRange -Range $workSheet1.Cells["A1:$($titleColumnEnd)"] -BackgroundColor LightBlue
$myExcelReport.Workbook.Worksheets[1].cells["A1:$($titleColumnEnd)"].Merge = $true
$myExcelReport.Workbook.Worksheets[1].cells["A1:$($titleColumnEnd)"].Style.HorizontalAlignment = "Center"
Close-ExcelPackage $myExcelReport 



