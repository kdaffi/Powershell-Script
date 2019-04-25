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

try{
$users = import-csv -Header Name -Delimiter ";" -Path $a
echo "`nReading file ......"

}
catch {[System.Exception]"Cannot Read / Create file !`n"}
foreach($test in $users)
{
    $count3 = 0;
    $found = 0;
    Write-Host "`n" $test.Name "`n" -ForegroundColor Green
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
    foreach($Domain in ($Domains))
    {
        $ADsPath = [ADSI]"LDAP://$Domain"
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	    $objSearcher.Filter = "(&(objectClass=group)(cn=$($test.Name)))"
	    $objSearcher.SearchScope = "Subtree"
 
	    $colResults = $objSearcher.FindAll()

        if ($count3 -eq 9 -and $found -eq 0){
                echo "Not Found"
        }

        foreach ($objResult in $colResults)
	    {
            $found = 1
            $DLGroup = $objResult.GetDirectoryEntry()

            $grpEnt = New-Object System.DirectoryServices.DirectoryEntry("LDAP://"+$DLGroup.distinguishedname)

            echo $Domain
            
            $ADsPath2 = [ADSI]$ADsPath.Path

            $ADsPath2.Children.Remove("GX SITI ALL HNS")
        }
    }
}