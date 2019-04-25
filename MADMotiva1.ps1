#################################################################

[environment]::CurrentDirectory
$inputFile = import-csv $ScriptDir\PCRefreshList.csv -Header Name -Delimiter ";"

#################################################################

try{
$ScriptDir = "C:\Users\DRA-CBJS08526-S\Desktop\PCScan1"
$users = import-csv $ScriptDir\PCRefreshList.csv -Header Name -Delimiter ";"
echo "`nReading file ......"
$file = New-Item "$env:userprofile\Desktop\Report ($(Get-Date -Format yyy-mm-dd-hhmm)).csv" -type file -Force -value "Computer Name,Group,Status,Date,Time`n"
}
catch {[System.Exception]"`nPCRefreshList.csv file NOT FOUND !`n"}


$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainList = @($objForest.Domains | Select-Object Name)
$Domains = $DomainList | foreach {$_.Name}
foreach($Domain in ($Domains))
{
    $ADsPath = [ADSI]"LDAP://$Domain"
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
    $objSearcher.Filter = "(&(objectClass=person)(cn=USLRIU))"
    $objSearcher.SearchScope = "Subtree"
 
    $colResults = $objSearcher.FindAll()
        
    foreach ($objResult in $colResults)
	{
        $Computer = $objResult.GetDirectoryEntry()
        $Computer.Properties["userprincipalname"].Value
    } 
}