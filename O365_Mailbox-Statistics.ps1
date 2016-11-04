# Connection to SommetEducation
Get-Credential $ga | Export-Clixml D:\PS-Scripts\sommetEducation_g-admin_Credential.xml

# Set this variable to the location of the file where credentials are cached
$credentialsCache = "D:\PS-Scripts\sommetEducation_g-admin_Credential.xml"

$o365Credentials = Import-Clixml $credentialsCache

Connect-MsolService -Credential $o365Credentials

Import-Module MSOnline
#Import-Module MSOnlineExtended

$o365DomainFilter="@sommet-education.com"
$outputFolder="d:\ps-scripts\"

cls

#Exchange Online
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $o365Credentials -Authentication "Basic" -AllowRedirection

Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber | Out-Null

# Get extensive information for multiple mailboxes
Get-mailbox | Select Alias,DisplayName, @{label="FirstName";expression={(Get-User -Identity $_.Name).FirstName}}, @{label="LastName";expression={(Get-User -Identity $_.Name).LastName}},SAMAccountName,Name, AddressListMembership,PrimarySmtpAddress,EmailAddresses,HiddenFromAddressListsEnabled,@{label="Title";expression={(Get-User -Identity $_.Name).Title}},@{label="Company";expression={(Get-User -Identity $_.Name).Company}},@{label="Department";expression={(Get-User -Identity $_.Name).Department}},Office,@{label="Phone";expression={(Get-User -Identity $_.Name).Phone}},@{label="MobilePhone";expression={(Get-User -Identity $_.Name).MobilePhone}} | Export-Csv c:\PS-Scripts\O365_SommetEducation_Mailboxes.csv 

# v1 (Return basic Mailbox statistics, with Display Name)
Get-Mailbox -ResultSize Unlimited | where {$_.primarysmtpaddress -like "*$o365DomainFilter"}| Get-MailboxStatistics | Select DisplayName, @{name=”Identity”; expression={$_.Identity}} ,@{name=”TotalItemSize (MB)”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}}, ItemCount |Sort “TotalItemSize (MB)” -Descending | Export-CSV “C:\PS-Scripts\Mailboxes_Statistics.csv” -NoTypeInformation -Encoding Default

# v2 (Return basic mailbox statistics, with PrimarySTMPAddress)
$(Foreach ($mailbox in Get-Mailbox -ResultSize Unlimited | Where {$_.primarysmtpaddress -like "*@$o365DomainFilter"})
{
    $stats = $mailbox | Get-MailboxStatistics | Select Identity, StorageLimitStatus,TotalItemSize,TotalDeletedItemSize,ItemCount,DeletedItemCount

	New-Object PSObject -Property @{
	    ObjectGUID = $Stats.Identity
	    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
	    Alias = $mailbox.Alias
	    DisplayName = $mailbox.DisplayName
	    FirstName = $mailbox.FirstName
	    LastName = $mailbox.LastName
	    TotalItemSize = $stats.TotalItemSize
	    TotalDeletedItemSize = $stats.TotalDeletedItemSize
	    ItemCount = $stats.ItemCount
        DeletedItemCount = $stats.DeletedItemCount
	}
}) | Select ObjectGUID,PrimarySMTPAddress,Alias,FirstName,LastName, DisplayName, HiddenFromAddressListsEnabled, AddressListMembership, TotalItemSize,TotalDeletedItemSize,ItemCount,DeletedItemCount | Export-CSV $("$outputFolder\{0}_{1}.csv" -f "Get-MailboxStatistics",$o365DomainFilter) -NTI -Encoding Default


#Clean-up > Remove Exchange Online session
#Remove-PSSession $exoSession
Get-PSSession | Remove-PSSession