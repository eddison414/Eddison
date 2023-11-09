<#
Please be advised that this script has been specifically developed for our unique IT infrastructure. It has been customized to align with the specific requirements and configurations of our system. Consequently, the script contains domain-specific information that has been adapted to suit our environment.

It is essential to understand that due to these customizations, the script is not designed for universal application. Attempting to implement it in a different IT environment may not yield the intended results and could potentially lead to system incompatibilities or operational issues.

We advise against using this script in any IT setting other than the one it was specifically created for. If you have any questions or require further clarification, please do not hesitate to reach out.
#>﻿﻿


# ----------------------------------[     Install required modules     ]-------------------------------------
Function required-Modules {
    <#
    This function has every module requiere for this Script to work
    
    #>
    
    $ExchangeOnline = Get-Module -Name ExchangeOnlineManagement
    if (!($ExchangeOnline )) { Install-Module -Name ExchangeOnlineManagement -confirm:$false }
    
    $PowerShellGet = Get-Module -Name PowerShellGet
    if (!($PowerShellGet)) { Install-Module -Name PowerShellGet -confirm:$false }
    
    $MicrosoftTeams = Get-Module -Name MicrosoftTeams
    if (!($MicrosoftTeams)) { Install-Module -Name MicrosoftTeams -confirm:$false } 
    
    $MSOnline = Get-Module -Name MSOnline
    if (!($MSOnline)) { Install-Module -Name MSOnline -confirm:$false }
    
    $AzureAD = Get-Module -Name AzureAD
    if (!($AzureAD)) { Install-Module -Name AzureAD -confirm:$false }
    
    $PnP = Get-Module PnP.PowerShell
    if (!($PnP)) { Install-Module -Name PnP.PowerShell -confirm:$false }
    
    $SqlServer = Get-Module -Name SqlServer
    if (!($SqlServer)) { Install-Module -Name SqlServer -confirm:$false }
    
    $iPilot = Get-Module -Name iPilot
    if (!($iPilot)) { Install-Module -Name iPilot -confirm:$false }
    
    $PSWriteColor = Get-Module -Name PSWriteColor
    if (!($PSWriteColor)) { Install-Module -Name PSWriteColor -confirm:$false }
    
    #KaceSMA
}

Function Encrypt {
    #### Set and encrypt our own password to file using default ConvertFrom-SecureString method
(get-credential).password | ConvertFrom-SecureString | set-content 'C:\Users\Automation\Documents\PowerShell Automation\E.txt'
}
    
# ----------------------------------[ Login into the require enviroment]-------------------------------

$encrypted = Get-Content 'C:\Users\Automation\Documents\PowerShell Automation\Encrypt.txt' | ConvertTo-SecureString
$prem = New-Object System.Management.Automation.PsCredential("automation", $encrypted)
$Cloud = New-Object System.Management.Automation.PsCredential("automation@ocdcrm.onmicrosoft.com", $encrypted)

<#
Important: this is a sample script. it was created to work on an spesifict IT enviroment. Please do not use in your inviroment. it will not work.
#>﻿


#Connect
Connect-PnPOnline -Url "ListT" -Credentials $Cloud
Initialize-iPilotSession -ApiKey 'auwpy2oQeK5KUYXvJAhid1ekgi8GrfesaC1qn0uy' -Credential $Cloud
Connect-MsolService -Credential $Cloud
Connect-ExchangeOnline -Credential $Cloud
Connect-AzureAD -Credential $Cloud
Connect-MicrosoftTeams -Credential $Cloud

#Test for connection making sure at least 1 of the DCs is online
$Connection = Test-Connection -ComputerName 'servername'
IF ($Connection) { $PSSessionDC1 = New-PSSession -ComputerName 'servername' -Credential $prem }
IF (!($Connection)) {
    $Connection = Test-Connection -ComputerName 'servername'
    IF ($Connection) { $PSSessionDC1 = New-PSSession -ComputerName 'servername' -Credential $prem }
}
IF (!($Connection)) { $PSSessionDC1 = New-PSSession -ComputerName 'servername' -Credential $prem }


$ichannel = New-PSSession -ComputerName 'servername' -Credential $prem



# ===================================== Termination ======================================
# Clear the screen and menu
CLS

Write-Host -ForegroundColor Green "*********************************************************************************************************"
Write-Host -ForegroundColor Red "  WARNING!!, Please be really really careful when running this script!   "
Write-Host -ForegroundColor Green "
*********************************************************************************************************



                                      Last Day termination Script

    This termination script will do the following for a terminated user;

    * Disable account.
    * Reset the user's password twice in the Active Directory. Refer to Generate Random Password.
    * Disable the user in Azure AD.
    * Revoke the user's Azure AD refresh tokens.
    * Hide from address book. 
    * Saved the AD groups in extensionAttribute1.
    * Remove from security and distro lists.
    * Convert to a shared mailbox.
    * Remove from nuwave / phone number.
    * Remove user from every teams group.
    * Move AD account to the Disabled OU.
    * Remove user from Domain user, and add them to Terminated Users.
    * Disable Ichannel account to get the license.

*********************************************************************************************************
"

Function Ichannel-disabled {

    param(
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [String]$cUserID # Ichannel login user ID
         )
    
    #Declare Servername
    
    #$sqlServer = "pkf-eip-sql1.odmd.local\ICHANNEL"
    #$database = "cadoc_system"
    #Invoke-sqlcmd Connection string parameters
    #$params = @{'server' = $sqlServer; 'Database' = $database }
    #$dataSet = new-object System.Data.Dataset
  
    # Create Ichannel account
    Invoke-Sqlcmd  -Query "spDeactivateInternalSubscriber 'ROOT', '$cUserID'" -As DataSet -ServerInstance "SQLINSTANCE" -Database "cadoc_system"
}

Function Disable-PKFODaCCOUNT{
    Param(

        [Parameter(Mandatory = $True,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [String]$Username
        
    )

    # Disable account
    Write-Host -ForegroundColor Yellow " * Disable account"
    Disable-ADAccount -Identity $Username

    # Reset the user's password twice in the Active Directory. Refer to Generate Random Password
    Write-Host -ForegroundColor Yellow " * Reset the user's password twice in the Active Directory. Refer to Generate Random Password"
    $Ns = 1..2 | foreach {
        $null = [Reflection.Assembly]::LoadWithPartialName("System.Web")
        $Password = [system.web.security.membership]::GeneratePassword(16, 2)
        Set-ADAccountPassword -Identity $Username -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force) }


    # Disable the user in Azure AD.
    Write-Host -ForegroundColor Yellow " * Disable the user in Azure AD."
    Set-AzureADUser -ObjectId "$Username@domain.com" -AccountEnabled $false

    # Revoke the user's Azure AD refresh tokens
    Write-Host -ForegroundColor Yellow " * Revoke the user's Azure AD refresh tokens"
    Revoke-AzureADUserAllRefreshToken -ObjectId "$Username@pkfod.com"

    # Hide from address book
    Write-Host -ForegroundColor Yellow " * Hide from address book"
    set-adobject -Identity ((Get-ADUser -Identity $Username).DistinguishedName) -replace @{msExchHideFromAddressLists = $true }
    Set-ADObject ((Get-ADUser -Identity $Username).DistinguishedName) -clear ShowinAddressBook

    # Disable Ichannel account
    Write-Host -ForegroundColor Yellow " * Disable Ichannel account"
    Ichannel-disabled -cUserID $Username


    # Remove from security and distro lists
    Write-Host -ForegroundColor Yellow " * Remove from security and distro lists"
    Get-ADUser -Identity $Username -Properties MemberOf | ForEach-Object {
    $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false }

    # Convert to a shared mailbox
    Write-Host -ForegroundColor Yellow " * Convert to a shared mailbox"
    Set-Mailbox -Identity $Username -Type Shared

    # Moved AD account to the Disabled OU
    Write-Host -ForegroundColor Yellow " * Move AD account to the Disabled OU"
    Move-ADObject -Identity ((Get-ADUser -Identity $Username).DistinguishedName) -TargetPath "OU=Disabled Group,OU=ODMD,DC=odmd,DC=local"

    # Remove user from every teams group
    Write-Host -ForegroundColor Yellow " * Remove user from every teams group (this part takes longer)..."
    Get-Team | foreach {
        try { Remove-TeamUser -GroupId $_.GroupId -User "$Username@domain.com" }
        catch {}} 

    # Remove Boston/woburm phones from user
    # Remove phone number
    $AzureADGroupMember = (Get-AzureADGroupMember -ObjectId '8ccc4f3f-beb7-45b2-9c79-4b5174ceef3b' -All $true | Where{$_.UserPrincipalName -like "$Username@domain.com"}).UserPrincipalName
    $Office = (get-aduser -Identity $Username -Properties office).office
   if(($AzureADGroupMember) -and (($Office -eq "MA - Woburn") -or ($Office -eq "MA - Boston") ))
   {Write-Host -ForegroundColor Yellow " * Remove phone number"
   $Null = Remove-iPilotTeamsUserAssignment -UserPrincipalName "$Username@pkfod.com" -iPilotDomain nuwms00076 -Credential $Cloud -ErrorAction SilentlyContinue}

   # Manual Steps
   $Steps = "
   * Sureprep/TaxCaddy 
   * Deactivate XCM Access 
   * Remove from Sharefile 
   * Remove from Netgain and QuickBooks online 
   * Netgain
   * Delete Adobe 
   * Mimecast
   * Engagement
        - Check local file room if possible  
   * Axcess Tax 
   * ClickSend
   * Remove U drive if they have it.
   "
               Send-MailMessage `
                –From 'automation@ocdcrm.onmicrosoft.com' `
                –To "servicedesk@domain.com" `
                –Subject "Termination Manual Steps for $Username" `
                –Body "$Steps" `
                -SmtpServer smtp.office365.com `
                -Credential $Cloud -UseSsl `
                -Port 587

   # Remove Office365 license after 5 days
   $RemoveLicense =  [PSCustomObject]@{
    Username = $Username
    TerminationDate = ((Get-Date).AddDays(5).ToString("MM/dd/yyyy"))
    Status = $Null
    }; $RemoveLicense | Export-Csv -Path "C:\Users\Automation\Documents\PowerShell Automation\Termination\LicenseRemoval.csv" -Append
}

Function Provide-mailboxaccess {
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory = $True)]
        [String]$Username 
    )


    # Ask question about wich manager will get access
    begin {
        Write-Host "
Give terminated user mailbox permission to one of the supervisors.

(1)	BAU/BTU        – Name Lastname
(2)	ABAS           – Name Lastname
(3)	PCG            – Name Lastname
(4)	BAG            – Name Lastname
(5)	Tax contro     – Name Lastname
(6) Finance        - Name Lastname
(7) HR             - Name Lastname
(8)        N/A
(9) ---   Others   ---


 An email will be send to the manager



" -ForegroundColor Green
        
        # Send email to manager
        Function Send-email {
            param ( 
                [Parameter(Mandatory = $True)]
                [String]$Terminatedinfo, 

                [Parameter(Mandatory = $True)]
                [String]$to 
            )

            $Body = "
    Hello,

    This email is to inform you that you now have access to the following user's Mailbox:

        $Terminatedinfo

    Please monitor activity and forward any relevant emails to the appropriate PKFOD. You should see the mailbox appear in the left side navigation in Outlook. 
    It may take a while for the Inbox contents to appear/finish sync'ing depending on the size of the mailbox, so please be patient.

    You will receive an email one day before IT permanently deletes the mailbox - typically 3 months to 6 months after termination date. 
    Please let us know if this should not be done for any reason.

    If you have any question, please send an email to servicedesk@pkfod.com.

        "

            Send-MailMessage `
                –From 'automation@ocdcrm.onmicrosoft.com' `
                –To "$to@domain.com" `
                –Subject "Please Monitor Offboarded User's Mailbox" `
                –Body "$Body " `
                -SmtpServer smtp.office365.com `
                -Credential $Cloud -UseSsl `
                -Port 587
        
        }
    }

    process {
        $Success = $False
        $TryCount = 0
        $MaxTries = 20
        while (-not ($Success) -and $TryCount -le $MaxTries) {
            try {

                # Sscript Block goes here
                [Validateset("1", "2", "3", "4", "5", "6", "7", "8", "9")][string]$Answer = Read-Host "Put a number from 1 to 8 base on the list bellow"
                $Success = $True
            }
            catch {
                # Error Logs / Scripts goes here
                Write-Host -ForegroundColor Red "Your selection has to be one of the numbers above, please try again"
            }
            $TryCount++
        }
    }


    end {

        Switch ($Answer) {
            "1" { $Manager = 'username' }
            "2" { $Manager = 'username' }
            "3" { $Manager = 'lrichards' }
            "4" { $Manager = 'username' }
            "5" { $Manager = 'username' }
            "6" { $Manager = 'username' }
            "7" { $Manager = 'username' }
            "8" { $Manager = 'N/A' }
            default { 
        
                $Success = $False; $TryCount = 0; $MaxTries = 20
                while (!($Success) -and $TryCount -le $MaxTries) {
                    try {

                        # Script Block goes here
                        # Ask for manager username
                        $Manager = Read-Host "If the user is not in the list above, please type the username"
                        # Cheack if the manager is in AD
                        $Manager = (Get-ADUser -Identity $Manager -ErrorAction SilentlyContinue).SamAccountName

                        # If the scrupt is Success, set to true
                        $Success = $True
                    }
                    catch {
                        # Error Logs / Scripts goes here
                        Write-Error "No user was found in AD with the username $Manager"
         
                    } 
                    $TryCount++
                }

            }
    
        }
        
        try {
            if (!($Manager -eq 'N/A')) {
                # Provide mailbox access
                $Null = Add-MailboxPermission  -Identity $Username -User $Manager -AccessRights FullAccess -AutoMapping $true -ErrorAction SilentlyContinue
                "* The mailbox for $Username has been given access to $Manager"
                # Send email to the manager
                Send-email -Terminatedinfo $Terminatedinfo -to $Manager
                $Note1 = "An email as been send to $to, for the mailbox access "
                $Global:Manager = $Manager
            }

        }
        Catch {}

    }
    
}

Function Account-confirmation {

    Write-Host "Please type the username for the terminated employee;"
    $Username = Read-Host "Username?"
    $ADUser = Get-ADUser -LDAPFilter "(sAMAccountName=$Username)" -Properties Title, Department


    $Name = $ADUser.Name
    $Email = $ADUser.UserPrincipalName
    $Title = $ADUser.Title
    $Department = $ADUser.Department
    $Terminatedinfo = "
    Name:        $Name
    Username:    $Username
    Email:       $Email
    Title:       $Title
    Department:  $Department
    "

    $answer = Read-Host "
Is this the user that has to be terminated?

$Terminatedinfo

*********************************************************************************************************
(yes/no)"


    if ($answer -eq "yes") {

        # Provide-mailboxaccess -Username $Username

        Write-Host -ForegroundColor Red " WARNING!!, this script will terminated $Name. This is the last change to stop" 
        $answer2 = Read-Host " Are you sure you want to continue? (yes/no)"

        if ($answer2 -eq "yes") { 

            Provide-mailboxaccess -Username $Username
            Disable-PKFODaCCOUNT -Username $Username; 

            $answer3 = Read-Host -ForegroundColor Green "Do you want to Terminated a another user?"
            if($answer3 -eq "Yes"){
            cls
            Account-confirmation
            }

        }
    }
}



# Run the script / F5
Account-confirmation