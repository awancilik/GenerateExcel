## ---------- Working with SQL Server ---------- ##
## https://powershell.org/forums/topic/using-a-poweshell-parameter-value-in-a-sql-query/
## https://www.comptia.org/blog/talk-tech-to-me-powershell-parameters-and-parameter-validation
## https://stackoverflow.com/questions/45758859/sql-extract-data-to-excel-using-powershell

## - Set Param:
[CmdletBinding()]
Param (
  [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
  [string]$StartDate,
  [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
  [string]$EndDate,
  [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
  [bool]$IsSendEmail,
  [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
  [string]$Env
)

## - Get SQL Server Table data:

$instance 	 = "" 
$database 	 = ""
$userid		 = "userid"
$password	 = "pwd"

if ($EndDate -eq "")
{
  $EndDate = $StartDate
}
	
## - Check PRD or Dev Env
if ($Env -eq "PRD")
{
	Write-Host "Report Log PRD" -Fore Red
	$instance="xx.xx.xx.xx" 
	$database="[Database Name]"
}
## - DEV Env
else
{
	Write-Host "Report Log DEV" -Fore Green
	$instance="xx.xx.xx.xx" 
	$database="[Database Name]"
}

$connString	="User ID=$userid;Password=$password;Initial Catalog=$database;Data Source=$instance"

try 
{
  ## - Connect to SQL Server using non-SMO class 'System.Data':
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection $connString
	$SqlConnection.Open()

	$sqlCommand = $SqlConnection.CreateCommand()
	$sqlCommand.CommandText = 
@"
	SELECT * FROM [Table Name] WHERE EntryDate BETWEEN '$StartDate' AND '$EndDate'
"@
	## - Extract and build the SQL data object '$DataSetTable':
	$adapter= New-Object System.Data.SqlClient.SqlDataAdapter $sqlCommand
	$dataset= New-Object System.Data.DataSet

	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCommand
	$DataSet = New-Object System.Data.DataSet

	$adapter.Fill($dataset) | Out-Null
	$DataSetTable = $dataset.Tables["Table"]

	## ---------- Working with Excel ---------- ##

	## - Create an Excel Application instance:
	$xlsObj = New-Object -ComObject Excel.Application;

	## - Create new Workbook and Sheet (Visible = 1 / 0 not visible)
	$xlsObj.Visible = 0;
	$xlsWb = $xlsobj.Workbooks.Add();
	$xlsSh = $xlsWb.Worksheets.item(1);

	## - Copy entire table to the clipboard as tab delimited CSV
	$DataSetTable | ConvertTo-Csv -NoType -Del "`t" | Clip

	## - Paste table to Excel
	$xlsObj.ActiveCell.PasteSpecial() | Out-Null

	## - Set columns to auto-fit width
	$xlsObj.ActiveSheet.UsedRange.Columns|%{$_.AutoFit()|Out-Null}

	## ---------- Saving file and Terminating Excel Application ---------- ##

	## - Saving Excel file - if the file exist do delete then save
	$fileName = Get-Date -Format "yyyy-MM-dd-HHmm"
	$reportFilename = "ReportLogUSUS_$fileName"
	$xlsFile = "[PathFolder]\$reportFilename.xls"
	$copyxlsFile = "[PathFolder]\SendEmail"
	
	## - Progress bar purpose
	for($i = 1; $i -lt 101; $i++ ) {for($j=0;$j -lt 10000;$j++) {} write-progress "Generate Excel Progress" "$i% Complete:" -perc $i;}

	if (Test-Path $xlsFile)
	{
		Remove-Item $xlsFile -Verbose
		$xlsObj.ActiveWorkbook.SaveAs($xlsFile);
		Copy-Item $xlsFile -Recurse -Destination $copyxlsFile -Verbose
	}
	else
	{
		$xlsObj.ActiveWorkbook.SaveAs($xlsFile);
		Copy-Item $xlsFile -Recurse -Destination $copyxlsFile -Verbose
	};

	## Quit Excel and Terminate Excel Application process:
	$xlsObj.Quit(); (Get-Process Excel*) | foreach ($_) { $_.kill() };
	
	## Send Email:
	if ($IsSendEmail)
	{			
	   $Date 		= get-date
       
	   $From 		= "[Email From]"
	   $To 			= "[Email To]"
	   $Subject 	= "Report Log - $Date"
	   $Body 		= "Please find attached"
	   $FileAttach 	= "$copyxlsFile\$reportFilename.xls"
	   $SMTPServer 	= "[SMTP]"
       
	   $Attachment 	= new-object Net.Mail.Attachment($FileAttach)
	   $SMTP 		= new-object Net.Mail.SmtpClient($SMTPServer)
	   $MSG 		= new-object Net.Mail.MailMessage($From, $To, $Subject, $Body)
	   $MSG.attachments.add($Attachment)
	   
	   $SMTP.send($msg)	  
	   
	   Write-Host "Success & Email Sent" -Fore Green
	}
	else
	{
	   Write-Host "Success"  -Fore green
	}
	
	Write-Host "URL Document = $xlsFile"
}
catch
{
	Write-Host $_ -Fore Red
}

## - End of Script - ##