
#################################### CONFIGURATION ####################################

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$CSVSourceDir = $scriptPath + '\DLCU_SC3\INPUT_FILE'  # Input file directory
$CSVLogMsgDir = $scriptPath + '\DLCU_SC3\DLC_LOGMESSAGE' # Log Message directory
$CSVReportDir = $scriptPath + '\DLCU_SC3_REPORT' # Report directory
$CSVServerListDir = $scriptPath + '\DLCU_SC3\SERVER_LIST' # Server List directory

########################################################################################

$flag = 0

if(!(Test-Path -Path $CSVSourceDir )){
    New-Item -ItemType directory -Path $CSVSourceDir
    $flag = 1
}
if(!(Test-Path -Path $CSVLogMsgDir )){
    New-Item -ItemType directory -Path $CSVLogMsgDir
}
if(!(Test-Path -Path $CSVReportDir )){
    New-Item -ItemType directory -Path $CSVReportDir
}
if(!(Test-Path -Path $CSVServerListDir )){
    New-Item -ItemType directory -Path $CSVServerListDir
    $flag = 1
}
if($flag -eq 1)
{
	Write-Host "`n`nNew input folder created automatically. Please put input file into the folder listed below`n`nFolder Name: INPUT_FILE`nFolder Path: $CSVSourceDir`n`nFolder Name: SERVER_LIST`nFolder Path: $CSVServerListDir`n`n" -ForegroundColor Red
	Write-Host "Press <enter> when the input file is ready`n`n" -ForegroundColor Green
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
while (!(Test-Path $CSVSourceDir\*.CSV))
{
	Write-Host "`nNo csv file available in the INPUT_FILE folder, Please put input file into the INPUT_FILE folder`nFolder Path: $CSVSourceDir`n" -ForegroundColor Red
	Write-Host "Press <enter> when the input file is ready`n`n" -ForegroundColor Green
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
while (!(Test-Path $CSVServerListDir\*.CSV))
{
	Write-Host "`nNo csv file available in the SERVER_LIST folder, Please put input file into the SERVER_LIST folder`nFolder Path: $CSVServerListDir`n" -ForegroundColor Red
	Write-Host "Press <enter> when the input file is ready`n`n" -ForegroundColor Green
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

$ServerListFiles = Get-ChildItem $CSVServerListDir -recurse -force | Where { $_.Name -like "*.csv" } | Foreach-Object -process { $_.FullName }
ForEach ($ServerFilesItem in $ServerListFiles)
{
    $ServerFileInfo = Get-Item $ServerFilesItem
    $ServerListCSVData = Import-CSV $ServerFilesItem -Delimiter "," -Header ServerName
}

#Delete old log files in directory
Remove-Item $CSVLogMsgDir\* -Include *.csv

try{
    Write-Output $ServerListCSVData
    Write-Host "`n-> Getting Message Tracking Log from all of the listed servers. It may take longer time, Please Wait ...`n" -ForegroundColor Green
    Invoke-Command { $ServerListCSVData | foreach { Get-MessageTrackingLog -server $_.ServerName -EventId Expand -ErrorAction 'silentlycontinue' -ResultSize unlimited | Select-Object RelatedRecipientAddress } | Export-Csv $CSVLogMsgDir\TEMP_DLC_LOGMESSAGE_$(Get-Date -Format yyyMMdd-HHmm).csv -NoTypeInformation }
    Write-Host "-> Status: SUCCESS`n" -BackgroundColor Black -ForegroundColor Green
}
catch{
    [system.exception]"Status: ERROR ! Extracting Message Tracking Log cannot be perform.`n"
}

###################### COMPARING DATA FROM INPUT FILES WITH MESSAGE LOG FILE ######################

Write-Host "-> Comparing data from input files with message log file. It may take longer time, Please Wait ...`n" -ForegroundColor Green

# Read all csv file in Directory
$DataFiles = Get-ChildItem $CSVSourceDir -recurse -force | Where { $_.Name -like "*.csv" } | Foreach-Object -process { $_.FullName }
$LogDataFiles0 = Get-ChildItem $CSVLogMsgDir -recurse -force | Where { $_.Name -like "*.csv" } | Foreach-Object -process { $_.FullName }


# Create new csv file for Report
$file = New-Item "$CSVReportDir\DLC_SC3_REPORT_$(Get-Date -Format yyyMMdd-HHmm).csv" -type file -Force -value "DistinguishedName,DisplayName,Alias,Owner,OwnerOU,OwnerCount,NoteField,PrimarySmtpAddress,WhenCreated,WhenChanged,HiddenfromGAL,MemberCount(Users),MemberCount(Groups),RecipientType,OrganizationalUnit,DLSentCount,Status`r`n"
$fileFullName = $file.FullName


ForEach($Files in $LogDataFiles0){
    $FileInfo = Get-Item $Files
    $a = Import-CSV $Files | Group-Object -Property RelatedRecipientAddress -NoElement
    Remove-Item $CSVLogMsgDir\* -Include *.csv
    $a | Select-Object @{Name="RelatedRecipientAddress";Expression={$_.Name}}, Count | Export-Csv $CSVLogMsgDir\DLC_LOGMESSAGE_$(Get-Date -Format yyyMMdd-HHmm).csv -NoTypeInformation
}

$LogDataFiles = Get-ChildItem $CSVLogMsgDir -recurse -force | Where { $_.Name -like "*.csv" } | Foreach-Object -process { $_.FullName }
ForEach ($LogDataFilesItem in $LogDataFiles)
{
    $LogFileInfo = Get-Item $LogDataFilesItem
    $CSVLogData = Import-CSV $LogDataFilesItem
}


ForEach ($DataFilesItem in $DataFiles)
    {
        $FileInfo = Get-Item $DataFilesItem
        $CSVData = Import-CSV $DataFilesItem -Delimiter ","
        foreach($test in $CSVData){
            try{
			foreach($Logtest in $CSVLogData){
                           if($test.PrimarySmtpAddress -eq $Logtest.RelatedRecipientAddress)
                           {
                                $status = "Active"
				$DLcount = $Logtest.Count
                                break;
                           }
                           else
                           {
                                $status = "Not Active"
				$DLcount = 0
                           }
                        }
                     
                }
                catch{[system.exception]"Status: ERROR ! Comparing files cannot be perform.`n"}

                # INSERT DATA INTO REPORT FILE
                try{
                    $dn = $test.DistinguishedName
                    $nf = $test.NoteField
                    $nf2 = $nf -replace "`"","`'"
                    $DispName = $test.DisplayName
                    $al = $test.Alias
                    $to = $test.Owner
                    $toOU = $test.OwnerOU
                    $OrgUnit = $test.OrganizationalUnit
                    $RecType = $test.RecipientType
                    $PrimarySMTP = $test.PrimarySmtpAddress
                    $wr = Write-Output ("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}" -f "`"$dn`"","`"$DispName`"","`"$al`"","`"$to`"","`"$toOU`"",$test.OwnerCount,"`"$nf2`"","`"$PrimarySMTP`"",$test.WhenCreated,$test.WhenChanged,$test.HiddenfromGAL,$test.'MemberCount(Users)',$test.'MemberCount(Groups)',"`"$RecType`"","`"$OrgUnit`"", $DLcount, $status)
                    $wr | Out-File $file -Append utf8
                }
                catch{[system.exception]"Status: ERROR ! Cannot append data into the output file.`n"}
        }
     }
     Write-Host "-> Status: SUCCESS`n" -BackgroundColor Black -ForegroundColor Green
     Write-Host "-> Report file has been created. The file can be found in $fileFullName`n" -ForegroundColor Green
     Write-Host "`n------------------------------------------- END -------------------------------------------`n" -ForegroundColor Red
     Read-Host -Prompt "`nPress <Enter> to exit"