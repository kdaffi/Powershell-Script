$ExchangeServer = "bejdc1-s-53401" # Exchange server name
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
Set-AdServerSettings -ViewEntireForest $true
Set-DistributionGroup -Id 'GX SITI EUC WS Front Office - Proactive' -HiddenFromAddressListsEnabled:$true
Set-DistributionGroup -Id 'GX-PT-DRA TestCreate 6' -HiddenFromAddressListsEnabled:$true
Remove-PSSession $Session