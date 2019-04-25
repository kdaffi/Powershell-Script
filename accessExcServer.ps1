

$ExchangeServer = "bejdc1-s-53401.asia-pac.shell.com"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
Import-PSSession $Session
Remove-PSSession $Session