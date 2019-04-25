$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainList = @($objForest.Domains | Select-Object Name)
$Domains = $DomainList | foreach {$_.Name}
foreach($Domain in ($Domains))
{
    #echo $Domain
        $ADsPath = [ADSI]"LDAP://$Domain"
	    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	    #$objSearcher.Filter = "(&(objectClass=group)(displayname=GX DS Nordic GSAP PGS))"
        $objSearcher.Filter = "(&(objectClass=person)(cn=USLRIU))"
	    $objSearcher.SearchScope = "Subtree"
 
	    $colResults = $objSearcher.FindAll()
        
            foreach ($objResult in $colResults)
	        {
                $Computer = $objResult.GetDirectoryEntry()
                # ConvertADSLargeInteger $Computer.pwdlastset.value
                #$aa =[datetime]::FromFileTimeUtc($Computer.properties.pwdlastset)

                $Date = [datetime]::FromFileTime($Computer.pwdlastset[0])
                
                echo $Date
                
            } 
}