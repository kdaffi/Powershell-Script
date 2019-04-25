$file = New-Item "$env:userprofile\Desktop\Report2 ($(Get-Date -Format yyy-mm-dd-hhmm)).csv" -type file -Force -value "DN, OU`n"
$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainList = @($objForest.Domains | Select-Object Name)
$Domains = $DomainList | foreach {$_.Name}
foreach($Domain in ($Domains))
{
        $ADsPath = [ADSI]"LDAP://$Domain"
	    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	    $objSearcher.Filter = "(&(objectClass=organizationalUnit)(name=Developers))"
	    $objSearcher.SearchScope = "Subtree"
 
	    $colResults = $objSearcher.FindAll()
        
        foreach ($objResult in $colResults)
	    {
            $Computer = $objResult.GetDirectoryEntry()
            $a = $Computer.Properties["distinguishedname"].Value
            $b = dsquery computer $a -limit 0
            foreach($c in $b)
            {
                write-host $c
                try{
                $wr = Write-Output ("{0},{1}" -f $c, "`"$a`"")
                $wr | Out-File $file -Append
                }
                catch{
                    [system.exception]"Group: $($n)`nStatus: Cannot append data into the output file !`n"
                }
            }
        } 
}



