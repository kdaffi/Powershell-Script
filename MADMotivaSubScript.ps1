Get-Recipient MYSDR7 | Select -ExpandProperty EmailAddresses | Where-Object { $_.IsPrimaryAddress -like 'False' } | Select SmtpAddress

Get-Recipient -resultsize "Unlimited" | Where { $_.emailaddresses -match "" } | select emailaddresses

Set-Mailbox "Dan Jump" -EmailAddresses SMTP:dan.jump@contoso.com,dan.jump@northamerica.contoso.com,danj@tailspintoys.com

Get-Mailbox MYSDR7 | Get-User | Select-Object firstname, lastname

Get-Recipient -resultsize "Unlimited" | Select -ExpandProperty EmailAddresses | Select SmtpAddress | Out-File C:\Users\K.Mohammed-SVR-A\Desktop\testAllEmail.csv

