$ErrorActionPreference = 'silentlycontinue'
try{
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
}

try{
$users = import-csv -Header Name -Delimiter ";" -Path $a
echo "`nReading file ......"
$file = New-Item "$env:userprofile\Desktop\GroupAvailabilityReport ($(Get-Date -Format yyy-mm-dd-hhmm)).csv" -type file -Force -value "Group Name,Status,Date,Time`n"
}
catch {[System.Exception]"Cannot Read / Create file !`n"}

$alluserdata = @()
$countDL = 0;
foreach($test in $users)
{
  if($test.'Name'){
    $count = 0
    $flag = 0
    $countDL += 1;
    Write-Host "`n" $countDL ") " $test.Name -ForegroundColor Green

    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
 
    foreach($Domain in ($Domains))
    {
        $count++;
	    $ADsPath = [ADSI]"LDAP://$Domain"
	    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	    $objSearcher.Filter = "(&(objectClass=group)(cn=$($test.Name)))"
	    $objSearcher.SearchScope = "Subtree"
 
	    $colResults = $objSearcher.FindAll()

        if ($colResults.Count -eq 0){
            $status = "Not Found"
        }
        else{
            $status = "Found"
            $count = 9
        }

        if($count -eq 9){
            echo $status
               try{
                    $wr = Write-Output ("{0},{1},{2},{3}" -f $test.name, $status, $(Get-Date -format d), $(Get-Date -format T))
                    $wr | Out-File $file -Append
                  }
               catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                   }
        }
       
       
    }
  }
}
Write-Host "------------------`nREPORT FILE OUTPUT`n------------------`n`nFilename: $($file.Name)`nFilepath: $($file.DirectoryName)" -ForegroundColor Red
$alluserdata | Out-GridView
Read-Host -Prompt "`nPress <Enter> to exit"