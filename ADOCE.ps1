$count = 0
$countMain = 0

function recurse{
    param([string]$DN)
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
    foreach($Domain in ($Domains))
    {
            $ADsPath = [ADSI]"LDAP://$Domain"
	        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	        $objSearcher.Filter = "(&(objectClass=person)(distinguishedname=" + $DN + "))"
	        $objSearcher.SearchScope = "Subtree"
 
	        $colResults = $objSearcher.FindAll()
        
            foreach ($objResult in $colResults)
	        {
                $mainUser = $objResult.GetDirectoryEntry()
                $count++
                if($mainUser.Properties["directreports"].Value)
                {
                    foreach ($items in $mainUser.Properties["directreports"].Value)
                        {
                            if ($items)
                            {                             
                                $child = GetChild -DN $items

                                $managerLevel = $count - 1

                                #Write-Host -NoNewline "`nUser Level: " $count "`nUser DN: " $child[5] "`nUser Name: " $child[4] "`nUser Mail: " $child[3] "`nManager Level: " $managerLevel "`nManager DN: " $child[2] "`nManager Name: " $child[0] "`nManager Mail: " $child[1] "`n"
                                if($count -eq 1)
                                {
                                    $countMain++
                                    Write-Host "`n" $countMain ") " $child[5]
                                }

                                $outputItems = @()
                                $outputItems = New-Object System.Object
                                $outputItems | Add-Member -MemberType NoteProperty -Name "User DN" -Value $child[5]
                                $outputItems | Add-Member -MemberType NoteProperty -Name "User Name" -Value $child[5]
                                $outputItems | Add-Member -MemberType NoteProperty -Name "User Mail" -Value $child[3]
                                $outputItems | Add-Member -MemberType NoteProperty -Name "User Level" -Value $count
                                $outputItems | Add-Member -MemberType NoteProperty -Name "Manager DN" -Value $child[2]
                                $outputItems | Add-Member -MemberType NoteProperty -Name "Manager Name" -Value $child[1]
                                $outputItems | Add-Member -MemberType NoteProperty -Name "Manager Mail" -Value $child[0]
                                $outputItems | Add-Member -MemberType NoteProperty -Name "Manager Level" -Value $managerLevel
                                $outputItems | Export-Csv -NoTypeInformation -Path "$env:userprofile\Desktop\Report.csv" -Append

                                recurse -DN $items
                            }
                        }
                        $count--
                }
                else
                {
                    $count--
                }
            } 
    }
}

function getChild{
    param([string]$DN)
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
    foreach($Domain in ($Domains))
    {
            $ADsPath = [ADSI]"LDAP://$Domain"
	        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	        $objSearcher.Filter = "(&(objectClass=person)(distinguishedname=" + $DN + "))"
	        $objSearcher.SearchScope = "Subtree"
 
	        $colResults = $objSearcher.FindAll()
        
            foreach ($objResult in $colResults)
	        {
                $childMail = "n/a";
                $Child = $objResult.GetDirectoryEntry()

                if($Child.Properties["userprincipalname"].Value)
                {
                    $childMail = $Child.Properties["userprincipalname"].Value
                }
                if($Child.Properties["mail"].Value)
                {
                    $childMail = $Child.Properties["mail"].Value
                }

                $childDisplayName = $Child.Properties["displayname"].Value

                if($Child.Properties["manager"].Value)
                {
                    $manager = getManager -ManagerDN $Child.Properties["manager"].Value
                }
            } 
    }
    return $manager[0], $manager[1], $manager[2], $childMail, $childDisplayName, $DN
}

function getManager{
    param([string]$ManagerDN)
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
    foreach($Domain in ($Domains))
    {
            $ADsPath = [ADSI]"LDAP://$Domain"
	        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	        $objSearcher.Filter = "(&(objectClass=person)(distinguishedname=" + $ManagerDN + "))"
	        $objSearcher.SearchScope = "Subtree"
 
	        $colResults = $objSearcher.FindAll()
        
            foreach ($objResult in $colResults)
	        {
                $managerMail = "n/a";
                $Manager = $objResult.GetDirectoryEntry()

                if($Manager.Properties["userprincipalname"].Value)
                {
                    $managerMail = $Manager.Properties["userprincipalname"].Value
                }
                if($Manager.Properties["mail"].Value)
                {
                    $managerMail = $Manager.Properties["mail"].Value
                }

                $managerDisplayName = $Manager.Properties["displayname"].Value
            } 
    }
    return $managerMail, $managerDisplayName, $ManagerDN
}

$inputDN = Read-Host -Prompt "`nPlease enter Distinguished Name"
#recurse -DN "CN=KLWGM1,OU=User Accounts Vista,OU=User Accounts,OU=User Directory,DC=asia-pac,DC=shell,DC=com"
recurse -DN $inputDN
