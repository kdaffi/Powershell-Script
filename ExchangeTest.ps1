$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainList = @($objForest.Domains | Select-Object Name)
$Domains = $DomainList | foreach {$_.Name}
 
foreach($Domain in ($Domains))
{
	$ADsPath = [ADSI]"LDAP://$Domain"
	$objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	$objSearcher.Filter = "(&(objectClass=group)(cn=GX-SITI-CBJ-ZZ-SITIEUCWSPKISupport))"
	$objSearcher.SearchScope = "Subtree"
 
	$colResults = $objSearcher.FindAll()
        
	foreach ($objResult in $colResults)
	{
        $DL = $objResult.GetDirectoryEntry()
        #Write-output $DL.msExchHideFromAddressLists

        $grpEnt = New-Object System.DirectoryServices.DirectoryEntry("LDAP://"+$DL.distinguishedName)
        $grpEnt.msExchHideFromAddressLists = "TRUE" 
        #$grpEnt.userAccountControl = "TRUE"
        $grpEnt.SetInfo()
        $grpEnt.Close()
    }
    }