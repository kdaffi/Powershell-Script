try
{
    $inputFile = import-csv $env:userprofile\Desktop\Input.csv -Header Name -Delimiter ";"
    echo "`nReading file ......"
    $file = New-Item "$env:userprofile\Desktop\Input2.csv" -type file -Force -value "DisplayName,CN`n"
}
catch
{
    [System.Exception]"`nInput.csv file NOT FOUND !`n"
}
$count = 0
foreach($test in $inputFile)
{
    $count++
    $found = 0
    Write-Host "`n" $count ") Group: " $test.Name -ForegroundColor Green
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
    foreach($Domain in ($Domains))
    {
        
        $ADsPath = [ADSI]"LDAP://$Domain"
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	    $objSearcher.Filter = "(&(objectClass=group)(displayname=$($test.Name)))"
	    $objSearcher.SearchScope = "Subtree"

	    $colResults = $objSearcher.FindAll()
        if($colResults.Count -gt 0){
            foreach ($objResult in $colResults)
	        {
                try
                {
                    $found = 1
                   [String]$Status = $objResult.Properties["cn"]
                }
                catch
                {
                    $Status = $_.Exception.Message
                }
                
            }
        }
    }

    if($found -eq 0)
    {
        $Status = "Not Found"
    }

    Write-Host "Status: " $Status

    try
    {
        $wr = Write-Output ("{0},{1}" -f $test.Name, $Status)
        $wr | Out-File $file -Append
    }
    catch
    {
        [system.exception]"Status: Cannot append data into the output file !`n"
    }
}