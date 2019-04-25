$total = 0
$file = New-Item "$env:userprofile\Desktop\AllShellEmail_($(Get-Date -Format yyy-mm-dd-hhmm)).csv" -type file -Force -value "MAIL, CN, DN`n"
$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainList = @($objForest.Domains | Select-Object Name)
$Domains = $DomainList | foreach {$_.Name}
foreach($Domain in ($Domains))
{
    $ADsPath = [ADSI]"LDAP://$Domain"
	$objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	$objSearcher.Filter = "(&(objectClass=user))"
	$objSearcher.SearchScope = "Subtree"
 
	$colResults = $objSearcher.FindAll()
        
    foreach ($objResult in $colResults)
	{
        $total++
        $getMail = $objResult.GetDirectoryEntry()
        $a = $getMail.Properties["mail"].Value
        if($a)
        {
            Write-Host -NoNewline $total ") " $a "`n"
            $b = $getMail.Properties["cn"].Value
            $c = $getMail.Properties["distinguishedname"].Value
            try{
                $wr = Write-Output ("{0},{1},{2}" -f "`"$a`"", "`"$b`"", "`"$c`"")
                $wr | Out-File $file -Append
               }
            catch{
                [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
               }
        }
        else
        {
            Write-Host -NoNewline $total ") No Mail ! `n"
        }
    } 
}
