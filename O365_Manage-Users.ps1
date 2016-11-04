################################################################################################################################################################  
# Script accepts 4 optional parameters from the command line  
#
# Arguments 
# InputFile - Optional - If needed this is used to target only recipients part of the designated domain  
# ConfigureLicense - Optional - G-Administrator Username ID for the tenant we are querying
# ResourceTypeFilter - Optional - Administrator Username password for the tenant we are querying  
#  
# To run the script
#  
# .\O365_Manage-Users.ps1 -InputFile D:\PS-Scripts\MasterData_Accounts.csv [-ConfigureLicense:$true]

################################################################################################################################################################  
  

#Accept input parameters  
Param(  

    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]  
    [string] $InputFile,
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]  
    [bool] $ConfigureLicense,
    [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $ResourceTypeFilter

)

#$InputFile = "D:\PS-Scripts\Sommet-Education\LesRoches.EDU_Mailboxes_Active.csv"

# InputFile value provided ?
if ([string]::IsNullOrEmpty($InputFile) -eq $true)
{ 
    Write-Error "Input file not specified, execution stopped!"
}
else {
    Write-Information "Input set to file $InputFile"
}

$msolcred = get-credential

connect-msolservice -credential $msolcred

#Set-ExecutionPolicy RemoteSigned

#Configure a new Exchange Online
$exoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $msolcred -Authentication Basic -AllowRedirection

#$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection

#Import an existing session (if exists)
Import-PSSession $exoSession -AllowClobber | Out-Null           

Write-Host "Retrieving Tenant details..." -ForegroundColor Yellow

$MSOLAccountSKU = Get-MsolAccountSku

if ($MSOLAccountSKU -ne $null)
{
    # in a multi-tenant environement 
    $tenantId = $MSOLAccountSKU.AccountObjectId[0]
    $tenantName = $MSOLAccountSKU.AccountName[0]

    Write-Host "Running on $tenantName | $tenantId"
}
  
# WARNING - Completely removes ALL users previously deleted
#Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin -Force

#Get-MsolAccountSku
#Get-MSolUser -All | Set-MsolUser -State $null -Phone $null -Fax $null -MobilePhone $null
   
$licEMSOptions = New-MsolLicenseOptions –AccountSkuId "sommeteducation:EMS"
$licOffProPlusFacultyOptions = New-MsolLicenseOptions –AccountSkuId "sommeteducation:OFFICESUBSCRIPTION_FACULTY"
$licStdFacultyOptions = New-MsolLicenseOptions –AccountSkuId "sommeteducation:STANDARDWOFFPACK_FACULTY" -DisabledPlans "OFFICE_FORMS_PLAN_2,PROJECTWORKMANAGEMENT,SWAY,YAMMER_EDU" #"RMS_S_ENTERPRISE,OFFICE_FORMS_PLAN_2,PROJECTWORKMANAGEMENT,SWAY,INTUNE_O365,YAMMER_EDU,SHAREPOINTWAC_EDU,MCOSTANDARD,SHAREPOINTSTANDARD_EDU,EXCHANGE_S_STANDARD"

#cls

$idxUserCount_New = 0
$idxUserCount_Update = 0
$idxUserCount_Skip = 0

Import-Csv $InputFile -Encoding Default -Delimiter ";" | ForEach {

    $msolUser = $null

    Write-Host "Processing $($_.ResourceType) | $($_.UserPrincipalName)" -ForegroundColor Gray

                #try geting a refence to the account, and store return message in the an error variable
                $msolUser = Get-MsolUser -UserPrincipalName $_.UserPrincipalName -ea SilentlyContinue -ev $errUserDoesNotExist

                if ($msolUser -eq $null)
                {

                    Write-Host "Account $($_.UserPrincipalName) not found, creating ...." -ForegroundColor Cyan
        
                    New-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation $_.UsageLocation -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -Title $_.JobTitle -Department $_.Department -Office $_.Office
                    
                    if($ConfigureLicense)
                    {
     
                        $errAddLicense = $null   
                    
                        Write-Host "Assigning licenses sommeteducation:STANDARDWOFFPACK_FACULTY to account $($_.UserPrincipalName)" -ForegroundColor Cyan
                    
                        #For new Users configure licenses plans too
                    
                        # License plan EMS
                        #Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses sommeteducation:EMS -LicenseOptions $licEMSOptions  -EA SilentlyContinue -ErrorVariable $errAddLicense

                        #if ($errAddLicense -ne $null) {
                        #    Write-Host "Error occured while adding license sommeteducation:EMS to $_.UserPrincipalName! Message $errAddLicense"
                        #    $errAddLicense = $null                    
                        #}

                        # License plan Office Pro-Plus
                        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses sommeteducation:OFFICESUBSCRIPTION_FACULTY -ea Continue -ev $errAddLicense

                        if ($errAddLicense -ne $null) {
                           Write-Host "Error occured while adding license sommeteducation:OFFICESUBSCRIPTION_FACULTY to $($_.UserPrincipalName)! Message $errAddLicense"
                           $errAddLicense = $null                    
                        }

                        # License plan Standard (SharePoint, Exchange, Skype for Business, etc.)
                        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses sommeteducation:STANDARDWOFFPACK_FACULTY -ea Continue -ev $errAddLicense
                    
                        if ($errAddLicense -ne $null) {
                            Write-Host "Error occured while adding license sommeteducation:STANDARDWOFFPACK_FACULTY to $($_.UserPrincipalName). Message $errAddLicense"
                            $errAddLicense = $null                    
                        }

                    }
                    else { Write-Information "Skipping license configuration as not specified" }

                    $idxUserCount_New ++
                }
                else
                {
                    Write-Host "Account located for $($msolUser.UserPrincipalName), updating...." -ForegroundColor Yellow

                    #Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation $_.UsageLocation -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -Title $_.JobTitle -Country $_.Country -Department $_.Department -Office $_.Office

                    Set-MsolUser -UserPrincipalName $($msolUser.UserPrincipalName) -UsageLocation $_.UsageLocation -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -Title $_.JobTitle -Country $_.Country -Department $_.Department -Office $_.Office
                    
                    if ($ConfigureLicense)
                    {
                       Write-Host "Is $($msolUser.UserPrincipalName) licensed ? $($msolUser.isLicensed)"
                       if ($msolUser.islicensed)
                       {
                            Write-Host "$($msolUser.UserPrincipalName) has $($msolUser.Licenses.AccountSkuId)"

                             # License plan Standard (Office ProPlus)
                            Set-MsolUserLicense -UserPrincipalName $msolUser.UserPrincipalName -LicenseOptions $licOffProPlusFacultyOptions -EA Continue -ErrorVariable $errAddLicense
                    
                            if ($errAddLicense -ne $null) {
                                Write-Host "Exception updating license sommeteducation:OFFICESUBSCRIPTION_FACULTY for $($msolUser.UserPrincipalName). Message $errAddLicense"
                                $errAddLicense = $null                 
                            }

                            # License plan Standard (SharePoint, Exchange, Skype for Business, etc.)
                            Set-MsolUserLicense -UserPrincipalName $msolUser.UserPrincipalName -LicenseOptions $licStdFacultyOptions -EA Continue -ErrorVariable $errAddLicense
                    
                            if ($errAddLicense -ne $null) {
                                Write-Host "Exception updating license sommeteducation:STANDARDWOFFPACK_FACULTY for $($msolUser.UserPrincipalName). Message $errAddLicense"
                                $errAddLicense = $null                 
                            }
                        }
                        else {
                
                               # User NOT licensed yed, adding NEW license plan Standard (SharePoint, Exchange, Skype for Business, etc.)
                            Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses sommeteducation:OFFICESUBSCRIPTION_FACULTY -LicenseOptions $licOffProPlusFacultyOptions -ea SilentlyContinue -ev $errAddLicense
                            Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses sommeteducation:STANDARDWOFFPACK_FACULTY -LicenseOptions $licStdFacultyOptions -ea SilentlyContinue -ev $errAddLicense
                    
                            if ($errAddLicense -ne $null) {
                                Write-Host "Error occured while adding license sommeteducation:STANDARDWOFFPACK_FACULTY to $($msolUser.UserPrincipalName). Message $errAddLicense"
                                $errAddLicense = $null                   
                            }
                     
                        } 
                    }
                    else {
                        Write-Information "Skipping license update as not specified"
                    }

                    $idxUserCount_Update ++   
                }

            #Set-Mailbox -Identity $_.UserPrincipalName -EmailAddresses $_.Aliases
            
            #Write-Host "User Account updated, setting Forward rules"
            #Set-Mailbox -Identity $_.UserPrincipalName -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $_.ForwardingSMTPAddress
}

Write-Host "New users $idxUserCount_New | Updated accounts $idxUserCount_Update"

#Clean-up > Remove Exchange Online session
#Remove-PSSession $exoSession
Get-PSSession | Remove-PSSession