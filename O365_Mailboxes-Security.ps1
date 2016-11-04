<#==================================================================================================================================================
Program  : Retrieve Mailbox Delegation Permission
version  : 0.1
Function : Retrieve and export to CSV delegate permission for a mailbox
made by aleresche 
===================================================================================================================================================#>

#Checking if the user want to establish a PS session to exchange online 

$o365DomainFilter= "*sommet-education.com"
$outputFolder = "D:\PS-Scripts\TMP"

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    
Import-PSSession $Session -AllowClobber | Out-Null

write-host "Starting to Retrieve Mailboxes Access Permissions..."

#Full Perm
Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySMTPAddress -like "*$o365DomainFilter"}| Get-MailboxPermission | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like “NT AUTHORITY\SELF”) -and -not ($_.User -like ‘*Discovery Management*’) } | Select Identity,User,AccessRights |  Export-Csv -LiteralPath ($outputFolder, "Mailboxes_Permissions.csv" -join "\") -NoTypeInformation -Encoding Default

#Get-Mailbox | Where {$_.PrimarySMTPAddress -like "*$o365DomainFilter"}| Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -ne "DiscoverySearchMailbox*" -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights | Export-Csv -NoTypeInformation C:\export\mailboxpermissionssource.csv
write-host "Done !........"  

##SEND ON BEHALF
write-host "Starting to Retrieve Mailboxes SendOnBehalf Permissions..."
Get-Mailbox  -ResultSize Unlimited | Where {$_.PrimarySMTPAddress -like "*$o365DomainFilter" -and $_.GrantSendOnBehalfTo -ne $Null} |Select Alias,GrantSendOnBehalfTo |  Export-Csv -LiteralPath ($outputFolder, "Mailboxes_SendOnBehalf.csv" -join "\") -NoTypeInformation -Encoding Default
write-host "Done !........"

##SENDAS
write-host "Starting to Retrieve Mailboxes Send As Permissions..."
Get-RecipientPermission -ResultSize Unlimited | where {$_.Trustee -notlike "*NT AUTHORITY\*" -and $_.Trustee -ne "NULL SID"} | Select Identity,Trustee |  Export-Csv -LiteralPath ($outputFolder, "Mailboxes_SendAs.csv" -join "\") -NoTypeInformation -Encoding Default
write-host "Done !........"

##ACCEPT MSG ONLY FROM
write-host "Starting to Retrieve Mailboxes Accept Message only from permissions..."
Get-Mailbox -ResultSize Unlimited | Where-Object {$_.PrimarySMTPAddress -like "*$o365DomainFilter" -and $_.AcceptMessagesOnlyFromSendersOrMembers -ne $Null} | Select Alias, @{Name='AcceptMessagesOnlyFromSendersOrMembers';Expression={[string]::join(";", @($_.AcceptMessagesOnlyFromSendersOrMembers))}} |  Export-Csv -LiteralPath ($outputFolder, "Mailboxes_AcceptMessagesFrom.csv" -join "\") -NoTypeInformation -Encoding Default
write-host "Done !........"

##REJECT MSG FROM
write-host "Starting to Retrieve Mailboxes rejected email address permissions..."
Get-Mailbox  -ResultSize Unlimited | Where-Object {$_.PrimarySMTPAddress -like "*$o365DomainFilter" -and $_.RejectMessagesFromSendersOrMembers -ne $Null} | Select Alias, @{Name='RejectMessagesFromSendersOrMembers';Expression={[string]::join(";", @($_.RejectMessagesFromSendersOrMembers))}} |  Export-Csv -LiteralPath ($outputFolder, "Mailboxes_RejectMessagesFrom.csv" -join "\") -NoTypeInformation -Encoding Default
write-host "Done !........"
 
##DISTRIBUTION LIST ACCEPT ONLY FROM
write-host "Starting to Retrieve Mailboxes rejected email address permissions..."
Get-DistributionGroup  -ResultSize Unlimited | Where-Object {$_.AcceptMessagesOnlyFromSendersOrMembers -ne $Null} | Select Alias, @{Name='AcceptMessagesOnlyFromSendersOrMembers';Expression={[string]::join(";", @($_.AcceptMessagesOnlyFromSendersOrMembers))}} |  Export-Csv -LiteralPath ($outputFolder, "DistributionGroups_AcceptMessagesFrom.csv" -join "\") -NoTypeInformation -Encoding Default
write-host "Done !........"
 
##DISTRIBUTION LIST REJECT FROM
Get-DistributionGroup  -ResultSize Unlimited | Where-Object {$_.RejectMessagesFromSendersOrMembers -ne $Null} | Select Alias, @{Name='RejectMessagesFromSendersOrMembers';Expression={[string]::join(";", @($_.RejectMessagesFromSendersOrMembers))}} | Export-Csv -LiteralPath ($outputFolder, "DistributionGroups_RejectMessagesFrom.csv" -join "\") -NoTypeInformation -Encoding Default -NoTypeInformation -Encoding Default
write-host "Done !........"

Write-host "CSVs extraction completed please verify file content..."

Get-PSSession | Remove-PSSession