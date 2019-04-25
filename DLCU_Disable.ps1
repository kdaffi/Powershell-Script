Set-AdServerSettings -ViewEntireForest $true
try{
$ScriptDir = "$env:userprofile\Desktop\DLCU"
$users = import-csv $ScriptDir\DLCU_Input.csv -Header Name -Delimiter ";"
echo "`nReading file ......"
$file = New-Item "$env:userprofile\Desktop\DLCU\Report\DLCU_Report_($(Get-Date -Format yyy-MM-dd-hhmmss)).csv" -type file -Force -value "DisplayName,StatusDate,Status`n"
}
catch {[System.Exception]"`nPCRefreshList.csv file NOT FOUND !`n"}
 
foreach($test in $users)
{
    Write-Output $test.Name
    try
    {
       Disable-DistributionGroup -Identity $test.Name -Confirm:$false 
       try
       {
         $wr = Write-Output ("{0},{1},{2}" -f $test.Name, $(Get-Date -format d), "Disable")
         $wr | Out-File $file -Append
       }
       catch
       {
         [system.exception]"Status: Cannot append data into the output file !`n"
       }
    }
    catch
    {
       Write-Output $_.Exception.Message
       try
       {
         $wr = Write-Output ("{0},{1},{2}" -f $test.Name, $(Get-Date -format d), $_.Exception.Message)
         $wr | Out-File $file -Append
       }
       catch
       {
         [system.exception]"Status: Cannot append data into the output file !`n"
       }
    }
}