$ADsPath = [ADSI]"LDAP://CN=GX-BS-BRU-ZZ-BSandSIAllStaffBrussels,OU=non-Logical Distribution Groups,OU=Group Directory,DC=europe,DC=shell,DC=com"
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
$objSearcher.SearchScope = "Subtree"
$colResults = $objSearcher.FindAll()
foreach ($objResult in $colResults)
{
    $Computer = $objResult.GetDirectoryEntry()
    foreach($a in $Computer.member)
    {
        $ADsPath1 = [ADSI]"LDAP://"+$Computer.member
        $objSearcher1 = New-Object System.DirectoryServices.DirectorySearcher($ADsPath1)
        $objSearcher1.SearchScope = "Subtree"
        $colResults1 = $objSearcher1.FindAll()
        foreach($objResult1 in $colResults1)
        {
            $Computer1 = $objResult1.GetDirectoryEntry()
            if($Computer1.objectClass = "person")
            {
                Write-Host $Computer1.member
            }
            else
            {
                $ADsPath2 = [ADSI]"LDAP://"+$Computer1.member
                $objSearcher2 = New-Object System.DirectoryServices.DirectorySearcher($ADsPath2)
                $objSearcher2.SearchScope = "Subtree"
                $colResults2 = $objSearcher2.FindAll()
                foreach($objResult2 in $colResults2)
                {
                    $Computer2 = $objResult2.GetDirectoryEntry()
                    Write-Host $Computer2.member
                }
            }
        }
    }
}