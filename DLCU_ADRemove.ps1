try
{
    $ScriptDir = "$env:userprofile\Desktop\DLCU"
    $inputFile = import-csv $ScriptDir\DLCU_Input.csv -Header Name -Delimiter ";"
    echo "`nReading file ......"
    $file = New-Item "$env:userprofile\Desktop\DLCU\Report\DLCU_Report_($(Get-Date -Format yyy-MM-dd-hhmmss)).csv" -type file -Force -value "DisplayName,StatusDate,Status`n"
}
catch
{
    [System.Exception]"`nDLCU_Input.csv file NOT FOUND !`n"
}

foreach($test in $inputFile)
{
    $found = 0
    Write-Host "`nGroup: " $test.Name -ForegroundColor Green
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
        if($colResults.Count -gt 0){
            foreach ($objResult in $colResults)
	        {
                try
                {
                    $found = 1
                    Remove-ADGroup -Identity $test.Name -Server $Domain -Confirm:$false -ErrorAction Stop
                    $Status = "Deleted"
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
        $wr = Write-Output ("{0},{1},{2}" -f $test.Name, $(Get-Date -format d), $Status)
        $wr | Out-File $file -Append
    }
    catch
    {
        [system.exception]"Status: Cannot append data into the output file !`n"
    }
}