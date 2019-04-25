$ErrorActionPreference = 'silentlycontinue'
<#try{
Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "(*.csv*)| *.csv*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
}

$a = Get-FileName -initialDirectory "\\tsclient\C\Users"
}
catch
{
  do
 {
    $a = Read-Host "Enter .csv filepath here (.csv)"
    $ext = (Get-Item $a).Extension
    if($ext -notlike ".csv"){
    write-host "NOT CSV FILE, PLEASE CHOOSE FILE WITH .CSV EXTENSION !`n" -ForegroundColor Red}
 }
 while($ext -notlike ".csv")
}#>

try{
$ScriptDir = "C:\Users\DRA-CBJS08526-S\Desktop\PCScan1"
$users = import-csv $ScriptDir\PCRefreshList.csv -Header Name -Delimiter ";"
echo "`nReading file ......"
$file = New-Item "$env:userprofile\Desktop\Report ($(Get-Date -Format yyy-mm-dd-hhmm)).csv" -type file -Force -value "Computer Name,Group,Status,Date,Time`n"
}
catch {[System.Exception]"`nPCRefreshList.csv file NOT FOUND !`n"}

$alluserdata = @()

foreach($test in $users)
{
  if($test.'Name'){
   $count3 = 0;
   $found = 0;
    Write-Host "`n" $test.Name "`n" -ForegroundColor Green

    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
 
    foreach($Domain in ($Domains))
    {
        $count3++;
	    $ADsPath = [ADSI]"LDAP://$Domain"
	    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	    $objSearcher.Filter = "(&(objectClass=Computer)(cn=$($test.Name)))"
	    $objSearcher.SearchScope = "Subtree"
 
	    $colResults = $objSearcher.FindAll()

        if ($count3 -eq 9 -and $found -eq 0){
            $status = "Computer Not Found"
                echo "`nStatus: Computer Not Found`n"
                try{
                $wr = Write-Output ("{0},{1},{2},{3},{4}" -f $test.name,$status, $status, $(Get-Date -format d), $(Get-Date -format T))
                $wr | Out-File $file -Append
                }
                catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                }
        }
        
	    
	    foreach ($objResult in $colResults)
	    {
            $found = 1;
		    $Computer = $objResult.GetDirectoryEntry()
            if ($Computer.useraccountcontrol -eq 4130){
            if($Computer.memberof -ne $null){
            foreach ($d in $Computer.memberof){
            $user2 = New-Object System.DirectoryServices.DirectoryEntry("LDAP://"+$d)

            $n = $user2.Properties["cn"].Value
            if($n -match "Domain Computers" -or $n -match "CL-Software_Restricted_China-GS" -or $n -match "CL-Software_Restricted_Russia-GS" -or $n -match "CL-Software_Restricted_Kazakhstan-GS" -or $n -like '*DomainComputers*'){
                $status = "Skip"
                echo "Group: $($n)`nStatus: Skip`n"
                try{
                $wr = Write-Output ("{0},{1},{2},{3},{4}" -f $test.name,$n, $status, $(Get-Date -format d), $(Get-Date -format T))
                $wr | Out-File $file -Append
                }
                catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                }
            }
            else
            {
                try {
                $status = "Delete"
                $grpEnt = New-Object System.DirectoryServices.DirectoryEntry("LDAP://"+$d)
                <#$grpEnt.Properties["member"].Remove($Computer.distinguishedname.ToString())
                $grpEnt.SetInfo()
                $grpEnt.Close()#>
                echo "Group: $($n)`nStatus: Delete`n"
                try{
                    $wr = Write-Output ("{0},{1},{2},{3},{4}" -f $test.name,$n, $status, $(Get-Date -format d), $(Get-Date -format T))
                    $wr | Out-File $file -Append
                }
                catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                }
                }catch{$status = "Error: Cannot delete user from group"
                try{
                $wr = Write-Output ("{0},{1},{2},{3},{4}" -f $test.name,$n, $status, $(Get-Date -format d), $(Get-Date -format T))
                $wr | Out-File $file -Append
                }
                catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                }
                [system.exception]"Group: $($n)`nStatus: Cannot delete user from group, you don't have credential !`n"}
            }
            
          }
          }
          else{
                $status = "Skip(No Group)"
                echo "`nStatus: Skip(No Group)`n"
                try{
                $wr = Write-Output ("{0},{1},{2},{3},{4}" -f $test.name,$status, $status, $(Get-Date -format d), $(Get-Date -format T))
                $wr | Out-File $file -Append
                }
                catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                }
          }
          }
          else{
                $status = "Skip(Computer Enable)"
                echo "`nStatus: Skip(Computer Enable)`n"
                try{
                $wr = Write-Output ("{0},{1},{2},{3},{4}" -f $test.name,$status, $status, $(Get-Date -format d), $(Get-Date -format T))
                $wr | Out-File $file -Append
                }
                catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                }
          }
          
	    }
    }
  }
}
Write-Host "------------------`nREPORT FILE OUTPUT`n------------------`n`nFilename: $($file.Name)`nFilepath: $($file.DirectoryName)" -ForegroundColor Red
$alluserdata | Out-GridView
Read-Host -Prompt "`nPress <Enter> to exit"