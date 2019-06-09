#*******************************************************************************
#* ORSPDF.ps1
#* Read html file and convert it to PDF format.
#* Carlos Kassab
#* 2019-June-06
#******************************************************************************* 

# For script testing open a command window and run:
# ORSPDF.bat \OpenReportingSystem\SampleData\htmlSample.html \OpenReportingSystem\SampleData\htmlSample.pdf
# ORSPDF.bat \OpenReportingSystem\SampleData\tmp994987126.html \OpenReportingSystem\SampleData\tmp994987126.pdf

if( $args.count -lt 1 ) 
{
	Write-output "Usage: ORSPDF.bat htmlfile pdffile.pdf"
	exit
} else {
	$sourceFileName = $($args[0])
	$targetFileName = $($args[1])
}

# Delete possible old file.
Remove-Item $targetFileName -ErrorAction Ignore 

& /OpenReportingSystem/Utils/wkhtmltopdf/wkhtmltox/bin/wkhtmltopdf.exe --page-size A4 -q $sourceFileName $targetFileName

