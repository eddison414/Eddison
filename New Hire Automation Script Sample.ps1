﻿
<#
# Important Notice Regarding Script Usage

Please be advised that this script has been specifically developed for a 'SAMPLE' IT Enviroment. 
This script is not a generic solution and should be treated with caution.

## Risks of Misapplication

It is essential to understand that due to each company different IT customizations, the script is not designed for universal application. This script is just mean to be use as a sample.
Attempting to implement it in a different IT environment may:
- Not yield the intended results
- Potentially lead to system incompatibilities or operational issues
- Cause unintended data modifications or loss
- Create security vulnerabilities if not properly adapted
- Result in compliance violations if used in a regulated industry without proper vetting

## Strong Recommendation

I strongly advise against using this script in any IT setting. If you are considering adapting this script for your environment:

1. Thoroughly review and understand each component of the script
2. Identify all organization-specific elements that would need modification
3. Consult with your IT security team to ensure compliance with your security policies
4. Test extensively in a sandboxed environment before any production use
5. Consider rebuilding the script from scratch to ensure full compatibility with your systems
6. Engage with professional IT consultants if you lack the in-house expertise to safely adapt the script

## Clarification
Remember, while this script may serve as a valuable reference or starting point, it should not be viewed as a plug-and-play solution for other environments. The safety and integrity of your IT infrastructure should always be the primary concern when considering the use of any external scripts or tools.
#>﻿

<#
.SYNOPSIS
Automates the IT onboarding process for new hires by performing the following key actions:
- Retrieves new hire data from SharePoint
- Manages license allocation
- Creates and configures Active Directory accounts
- Synchronizes with Azure AD
- Sets up Exchange mailboxes
- Manages group memberships
- Creates and configures home folders
- Updates physical address information
- Assigns and configures phone numbers
- Updates SharePoint lists with onboarding progress
- Sends notification emails
- Handles errors and performs logging throughout the process

This script integrates with multiple systems including SharePoint, Active Directory, 
Exchange Online, and Azure AD to streamline the onboarding workflow.
#>

# ----------------------------------[     Install required modules     ]-------------------------------------
Function required-Modules {
    <#
    .SYNOPSIS
    This function checks for and installs required modules for the script.
    
    .DESCRIPTION
    It iterates through a list of required modules, checks if they are installed,
    and installs them if they are not present.
    #>
    
    $requiredModules = @(
        "ExchangeOnlineManagement",
        "PowerShellGet",
        "MicrosoftTeams",
        "MSOnline",
        "AzureAD",
        "PnP.PowerShell",
        "SqlServer",
        "PSWriteColor"
        # Add any other required modules here
    )

    foreach ($module in $requiredModules) {
        if (!(Get-Module -Name $module -ListAvailable)) {
            Write-Host "Installing module: $module"
            Install-Module -Name $module -Force -Confirm:$false
        }
        else {
            Write-Host "Module already installed: $module"
        }
    }
}
    
# ----------------------------------[ Login into the require enviroment]-------------------------------

<# 
The way we login into the different environments will change depending on the company. 
This section will need to be customized for each organization's specific needs
#>
    
# ----------------------------------[     Functions     ]-------------------------------------
  
Function AD-Account {
    <#
    .SYNOPSIS
    Creates a new Active Directory account for a new hire.

    .DESCRIPTION
    This function creates a new Active Directory account based on the provided new hire information.
    It sets various AD attributes and ensures the account is created with proper naming conventions and settings.

    .PARAMETER NewHire
    An object containing the new hire's information.
    #>
    param($NewHire)

    if ($NewHire.LicenseCheck -eq 'Pass') {
        # Construct the display name
        $DisplayName = if ($NewHire.Middle) { 
            "$($NewHire.FirstName) $($NewHire.Middle) $($NewHire.LastName)" 
        } else { 
            "$($NewHire.FirstName) $($NewHire.LastName)" 
        }

        # Generate a unique username
        $UserPrincipalName = Username-Check -firstName $NewHire.FirstName -LastName $NewHire.LastName
        $UserPrincipalName = $UserPrincipalName.ToLower()

        # Create the AD account
        $Success = Invoke-Command -Session $PSSessionDC1 -ScriptBlock {
            $Success = $False
            $TryCount = 0
            $MaxTries = 5

            while (!$Success -and $TryCount -le $MaxTries) {
                try {
                    # Create a new Active Directory Account 
                    New-ADUser -Name $args[0] `
                        -DisplayName $args[0] `
                        -SamAccountName $args[1] `
                        -Path $args[2] `
                        -UserPrincipalName $args[3] `
                        -AccountPassword $args[4] `
                        -EmailAddress $args[3] `
                        -GivenName $args[5] `
                        -Surname $args[6] `
                        -Office $args[7] `
                        -Enabled $True `
                        -Description $args[8] `
                        -Title $args[8] `
                        -Department $args[9] `
                        -Company "COMPANY_NAME" `
                        -HomePage "www.company.com" `
                        -EmployeeID $args[10]

                    $Success = $True
                }
                catch {
                    $TryCount++
                    Write-Error "Attempt $TryCount failed: $_"
                    Start-Sleep -Seconds 5
                }
            }

            $Success
        } -ArgumentList $DisplayName, 
                        $UserPrincipalName, 
                        (Get-OU -NewHire $NewHire), 
                        ($UserPrincipalName + "@yourdomain.com"),
                        (Create-Password -StartDate $NewHire.StartDate),
                        $NewHire.FirstName,
                        $NewHire.LastName,
                        (AD-Office -Location $NewHire.City),
                        $NewHire.JobTitle,
                        $NewHire.Department,
                        $NewHire.EmployeeID

        if ($Success -eq $True) {
            $NewHire | Add-Member -NotePropertyMembers @{
                'Username' = $UserPrincipalName
                'Email' = ($UserPrincipalName + '@yourdomain.com')
            }
            return $NewHire
        }
    }
}

function Manage-ADUserMobilePhone {
    Param(
        [string]$UserPrincipalName,
        [string]$MobilePhone,
        [String]$showNumber
    )


    if ($showNumber.ToLower() -eq 'yes') {
        # Add the mobile phone number to the AD account
        Set-ADUser -Identity $UserPrincipalName -MobilePhone $MobilePhone
    }
    elseif ($showNumber.ToLower() -eq 'no') {
        # Add the mobile phone number to the AD account
        Set-ADUser -Identity $UserPrincipalName -MobilePhone $MobilePhone -Server 'HA-DC1'

        # Define the scheduled task parameters
        $TaskName = "RemoveMobilePhone-$UserPrincipalName"
        $TaskStartTime = (Get-Date).AddDays(5)

        # Create the scheduled task to remove the mobile phone number after 5 days
        $ScheduledTaskAction = New-ScheduledTaskAction -Execute "powershell.exe" `
            -Argument "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -Command ""Set-ADUser -Identity '$UserPrincipalName' -Clear Mobile"""
        $Trigger = New-ScheduledTaskTrigger -Once -At $TaskStartTime
        $Settings = New-ScheduledTaskSettingsSet -DontStopOnIdleEnd

        $ScheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $ScheduledTaskAction -Trigger $Trigger -Settings $Settings

    }
    else {

    }
}

Function Set-LogonScript {
    <#
    .SYNOPSIS
    Assigns an appropriate logon script to a new user based on their department and location.

    .DESCRIPTION
    This function determines the appropriate logon script for a new user based on their department,
    job title, and office location. It then assigns this script to the user's AD account.

    .PARAMETER NewHire
    An object containing the new hire's information, including department, job title, and office location.
    #>
    Param($NewHire)

    # Initialize logon script variable
    $LogonScript = $null

    # Determine logon script based on department
    switch -regex ($NewHire.department) {
        "^IT" { $LogonScript = 'IT_LogonScript.bat' }
        "^HR" { $LogonScript = 'HR_LogonScript.bat' }
        "^Finance" { $LogonScript = 'Finance_LogonScript.bat' }
        # Add more department-specific scripts as needed
    }

    # If no department-specific script, check job title
    if (-not $LogonScript) {
        switch -regex ($NewHire.title) {
            "Manager" { $LogonScript = 'Manager_LogonScript.bat' }
            "Executive" { $LogonScript = 'Executive_LogonScript.bat' }
            # Add more title-specific scripts as needed
        }
    }

    # If still no script assigned, use location-based script
    if (-not $LogonScript) {
        switch ($NewHire.office) {
            "Headquarters" { $LogonScript = 'HQ_LogonScript.bat' }
            "Branch_Office" { $LogonScript = 'Branch_LogonScript.bat' }
            "Remote" { $LogonScript = 'Remote_LogonScript.bat' }
            # Add more location-specific scripts as needed
            default { $LogonScript = 'Default_LogonScript.bat' }
        }
    }

    # Assign the determined logon script to the user
    try {
        Set-ADUser -Identity $NewHire.Username -ScriptPath $LogonScript -ErrorAction Stop
        Write-Host "Logon script '$LogonScript' successfully assigned to $($NewHire.Username)"
    }
    catch {
        Write-Error "Failed to assign logon script to $($NewHire.Username): $_"
    }
}

function AD-Office {
    <#
    .DESCRIPTION
    This function takes a city name as input and returns a standardized office location string.
    It's used to ensure consistency in office designations across Active Directory entries.

    .PARAMETER Location
    The name of the city where the employee is located.

    .EXAMPLE
    AD-Office -Location "CityA"
    Returns: "State1 - CityA"

    .NOTES
    The function uses a switch statement to map cities to their corresponding state or region.
    Cities not explicitly listed will be assigned to the default HQ location.
    #>

    param(
        [Parameter(Mandatory = $True)]
        [AllowEmptyString()]
        [String]$Location
    )

    # Use switch statement to determine the appropriate office designation
    switch ($Location) {
        # Group cities by state or region
        { "CityA", "CityB", "CityC", "CityD" -contains $_ } { 
            $Office = "State1 - $Location" 
        }
        { "CityE", "CityF", "CityG" -contains $_ } { 
            $Office = "State2 - $Location" 
        }
        { "CityH", "CityI" -contains $_ } { 
            $Office = "State3 - $Location" 
        }
        { "CityJ", "CityK", "CityL" -contains $_ } { 
            $Office = "State4 - $Location" 
        }
        "CityM" { 
            $Office = "State5 - $Location" 
        }
        "CityN" { 
            $Office = "State6 - $Location" 
        }
        # International locations
        { "CityO", "CityP", "CityQ" -contains $_ } { 
            $Office = "International" 
        }
        # Default case for any unspecified locations
        default { 
            $Office = "HQ - DefaultCity" 
        }
    }

    # Return the standardized office designation
    $Office
}


Function Waitfor-sync {
    <#
    .DESCRIPTION
    This function attempts to add a user to specified Azure AD groups based on their status.
    It retries the operation multiple times if it fails, with increasing wait times between attempts.

    .PARAMETER Email
    The email address of the user to be added to the groups.

    .PARAMETER Username
    The username of the user to be added to the groups.

    .PARAMETER Status
    The status of the user, which determines which groups they will be added to.

    .NOTES
    This function assumes that Azure AD sync is running and may take some time to complete.
    If all attempts fail, it will trigger a cleanup process and log the failure.
    #>
    
    Param(
        [string]$Email,
        [string]$Username,
        [string]$Status
    )

    $Success = $False
    $TryCount = 0
    $MaxTries = 10
    $Waiting = 300  # Initial wait time in seconds

    while (!$Success -and $TryCount -le $MaxTries) {
        try {
            # Attempt to add user to a general exclusion group
            Add-AzureADGroupMember -ObjectId "GeneralExclusionGroupID" -RefObjectId (Get-AzureADUser -ObjectId $Email).ObjectId -ErrorAction Stop

            # Add user to specific groups based on their status
            if ($Status -like "Full Time Employee") {
                Add-AzureADGroupMember -ObjectId "FullTimeEmployeeGroupID" -RefObjectId (Get-AzureADUser -ObjectId $Email).ObjectId -ErrorAction Stop
                Write-Host "User $Username added to Full Time Employee group" -ForegroundColor Green
            }
            else {
                Add-AzureADGroupMember -ObjectId "OtherEmployeeGroupID" -RefObjectId (Get-AzureADUser -ObjectId $Email).ObjectId -ErrorAction Stop
                Write-Host "User $Username added to Other Employee group" -ForegroundColor Green
            }
    
            $Success = $True
        }
        catch {
            $TryCount++

            if ($TryCount -in @(2, 5, 8)) {
                # Trigger AD sync on specific attempts
                Invoke-Command -Session $PSSessionDC1 -ScriptBlock { Start-ADSyncSyncCycle } -ErrorAction SilentlyContinue
            }

            if ($TryCount -le 9) {
                Write-Host "Attempt $TryCount: User $Email not found in Azure AD. Waiting $($Waiting / 60) minutes before next attempt." -ForegroundColor Yellow
                Start-Sleep -Seconds $Waiting
                $Waiting += 300  # Increase wait time for next attempt
            }

            if ($TryCount -eq $MaxTries) {
                # Cleanup process if all attempts fail
                Remove-ADUser -Server "DC1" -Identity $Username -Confirm:$false
                Write-Host "Max attempts reached. AD account for $Username has been removed. Process will be retried later." -ForegroundColor Red
                
                # Log the failure
                $Global:SyncFailureLog += "Sync failed for user $Username after $MaxTries attempts."
            }
        }
    }
}


Function License-Count {
    <#
    .DESCRIPTION
    This function retrieves and calculates the number of available licenses for various
    Microsoft 365 products. It stores the results in global variables for later use.

    .NOTES
    This function assumes you have the necessary permissions to query license information
    and that you're already connected to the Microsoft 365 tenant.
    #>

    # Initialize counters for license requests
    $Global:Request_LicenseA = 0
    $Global:Request_LicenseB = 0
    $Global:Request_LicenseC = 0

    # Function to get available licenses for a specific product
    function Get-AvailableLicenses {
        param (
            [string]$ProductName,
            [string]$LicenseSku
        )
        $LicenseInfo = Get-MsolAccountSku | Where-Object { $_.AccountSkuId -like $LicenseSku }
        $AvailableLicenses = $LicenseInfo.ActiveUnits - $LicenseInfo.ConsumedUnits
        Write-Host "Available $ProductName licenses: $AvailableLicenses" -ForegroundColor Green
        return $AvailableLicenses
    }

    # Get license counts for different products
    $Global:LicenseA = Get-AvailableLicenses -ProductName "License A" -LicenseSku "TenantName:LicenseA"
    $Global:LicenseB = Get-AvailableLicenses -ProductName "License B" -LicenseSku "TenantName:LicenseB"
    $Global:LicenseC = Get-AvailableLicenses -ProductName "License C" -LicenseSku "TenantName:LicenseC"

    # Log the license counts
    Write-Host "License count retrieval complete." -ForegroundColor Yellow
}
         
Function License-check {
    <#
    .DESCRIPTION
    This function checks the availability of licenses for new hires and determines
    if there are enough licenses available based on the employee's status.

    .PARAMETER NewHireList
    An array of new hire objects containing information about each new employee.

    .NOTES
    This function assumes that global variables for license counts and requests
    have been initialized by a previous function call (e.g., License-Count).
    #>
    param($NewHireList)

    Begin {
        # Initialize any necessary variables
    }

    Process {
        foreach ($User in $NewHireList) {
            $FirstName = $User.FirstName
            $LicenseCheckPassed = $true

            # Check License A availability
            $Global:LicenseA = ($Global:LicenseA - 1)
            if ($Global:LicenseA -gt 1) {
                Write-Host "License A available for $FirstName $($User.LastName)" -ForegroundColor Green
            }
            else {
                Write-Host "License A not available for $FirstName $($User.LastName)" -ForegroundColor Red
                $Global:Request_LicenseA += 1
                $LicenseCheckPassed = $false
            }

            # Check License B availability for full-time employees
            if ($User.Status -eq "Full Time Employee") {
                $Global:LicenseB = ($Global:LicenseB - 1)
                if ($Global:LicenseB -gt 1) {
                    Write-Host "License B available for $FirstName $($User.LastName)" -ForegroundColor Green
                }
                else {
                    Write-Host "License B not available for $FirstName $($User.LastName)" -ForegroundColor Red
                    $Global:Request_LicenseB += 1
                    $LicenseCheckPassed = $false
                }
            }

            # Check License C availability
            $Global:LicenseC = ($Global:LicenseC - 1)
            if ($Global:LicenseC -gt 1) {
                Write-Host "License C available for $FirstName $($User.LastName)" -ForegroundColor Green
            }
            else {
                Write-Host "License C not available for $FirstName $($User.LastName)" -ForegroundColor Red
                $Global:Request_LicenseC += 1
                $LicenseCheckPassed = $false
            }

            # Determine if all required licenses are available
            if ($LicenseCheckPassed) {
                $requiredLicenses = if ($User.Status -eq "Full Time Employee") { "3/3" } else { "2/2" }
                Write-Host "User $FirstName $($User.LastName) has passed ($requiredLicenses) License check" -ForegroundColor Green
                $User | Add-Member -NotePropertyMembers @{'LicenseCheck' = "Pass" } -Force
            }
            else {
                Write-Host "Not enough licenses available for $FirstName $($User.LastName)" -ForegroundColor Yellow
            }

            # Return the updated user object
            $User
        }
    }

    End {
        # Perform any necessary cleanup or final operations
    }
}

Function Waitfor-Mailbox {
    <#
    .DESCRIPTION
    This function checks for the existence of a mailbox for a newly created user.
    It attempts to retrieve the mailbox multiple times, waiting between attempts.

    .PARAMETER Email
    The email address of the user whose mailbox is being checked.

    .PARAMETER Username
    The username of the user, used for logging and potential cleanup operations.

    .NOTES
    This function assumes you have the necessary permissions to query Exchange Online
    and that you're already connected to the Exchange Online service.
    #>
    [CmdletBinding()]
    Param(
        [string]$Email,
        [string]$Username
    )

    $Success = $False
    $TryCount = 0
    $MaxTries = 10
    $WaitTimeSeconds = 60

    while (!$Success -and $TryCount -lt $MaxTries) {
        try {
            $Mailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
            
            if ($Mailbox) { 
                Write-Host "Mailbox for $Email has been created successfully." -ForegroundColor Green
                $Success = $True
            }  
        }
        catch {
            $TryCount++
            
            if ($TryCount -eq $MaxTries) {
                Write-Host "Failed to find mailbox for $Email after $MaxTries attempts." -ForegroundColor Red
                Write-Host "Initiating cleanup process for $Username" -ForegroundColor Yellow
                
                # Placeholder for cleanup process
                # In a real scenario, you might want to remove the AD account or perform other cleanup tasks
                # Remove-ADUser -Identity $Username -Confirm:$false
                
                # Log the failure
                $Global:MailboxCreationFailures += "Failed to create mailbox for $Username after $MaxTries attempts."
            }
            else {
                Write-Host "Attempt $TryCount: Mailbox for $Email not found. Waiting $WaitTimeSeconds seconds before next attempt." -ForegroundColor Yellow
                Start-Sleep -Seconds $WaitTimeSeconds
            }
        }
    }

    if ($Success) {
        return $True
    }
    else {
        return $False
    }
}

Function Username-Check {
    <#
    .DESCRIPTION
    This function generates a unique username for a new user based on their first and last name.
    It checks for existing usernames in Active Directory and adjusts the username if conflicts are found.

    .PARAMETER firstName
    The first name of the new user.

    .PARAMETER lastName
    The last name of the new user.

    .NOTES
    This function assumes you have the necessary permissions to query Active Directory.
    It follows a specific pattern for username generation and conflict resolution.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$firstName,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$lastName
    )

    $domain = "@example.com"
    $email = $firstName[0] + $lastName + $domain

    # Function to check if a username exists in AD
    function Test-ADUsername {
        param([string]$Username)
        return [bool](Get-ADUser -Filter {SamAccountName -eq $Username} -ErrorAction SilentlyContinue)
    }

    # Check for unique email in AD
    if (Test-ADUsername ($email -replace $domain)) {
        $email = $firstName + '.' + $lastName + $domain
        
        if (Test-ADUsername ($email -replace $domain)) {
            # If both attempts fail, generate a unique username
            $counter = 1
            do {
                $email = $firstName[0] + $lastName + $counter + $domain
                $counter++
            } while (Test-ADUsername ($email -replace $domain))
        }
    }

    $username = ($email -split '@')[0]
    Write-Verbose "Generated username: $username"
    return $username
}


function change-physicaladdress {
    <#
    .DESCRIPTION
    This function updates the physical address attributes of a user in Active Directory
    based on their office location.

    .PARAMETER Username
    The username of the user whose address information needs to be updated.

    .NOTES
    This function assumes you have the necessary permissions to query and modify
    Active Directory user objects.
    #>
    param(
        [Parameter(Mandatory = $True)]
        [String]$Username
    )
    
    # Retrieve the user's office location from Active Directory
    $Location = (Get-ADUser -Identity $Username -Properties Office).Office

    # Define address information based on office location
    switch ($Location) {
        "Location1" {
            $AddressInfo = @{
                StreetAddress = "Street Address 1"
                City          = "City1"
                State         = "State1"
                PostalCode    = "PostalCode1"
                Country       = "Country1"
                OfficePhone   = "+1 000 000-0000"
            }
        }
        "Location2" {
            $AddressInfo = @{
                StreetAddress = "Street Address 2"
                City          = "City2"
                State         = "State2"
                PostalCode    = "PostalCode2"
                Country       = "Country2"
                OfficePhone   = "+1 000 000-0000"
            }
        }
        default {
            Write-Warning "Unknown office location: $Location. Using default address."
            $AddressInfo = @{
                StreetAddress = "Default Street Address"
                City          = "Default City"
                State         = "Default State"
                PostalCode    = "Default PostalCode"
                Country       = "Default Country"
                OfficePhone   = "+1 000 000-0000"
            }
        }
    }

    # Update AD attributes for the user
    try {
        Set-ADUser -Identity $Username -Replace @{
            StreetAddress = $AddressInfo.StreetAddress
            City          = $AddressInfo.City
            State         = $AddressInfo.State
            PostalCode    = $AddressInfo.PostalCode
            Country       = $AddressInfo.Country
        }
        Set-ADUser -Identity $Username -OfficePhone $AddressInfo.OfficePhone
        Write-Host "Successfully updated physical address for user: $Username" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to update physical address for user: $Username. Error: $_"
    }
}



Function MailboxSettings {
    <#
    .DESCRIPTION
    This function applies several settings and permissions to a specified mailbox in Exchange Online.

    .PARAMETER UserPrincipalName
    The user principal name (UPN) of the user whose mailbox settings need to be configured.

    .NOTES
    This function assumes you have the necessary permissions to modify Exchange Online mailboxes
    and that you're already connected to Exchange Online PowerShell.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [String]$UserPrincipalName
    )

    $MaxTries = 10
    $TryCount = 0

    do {
        try {
            # Add 'SendAs' permissions for the mailbox
            Add-RecipientPermission -Identity $UserPrincipalName -Trustee "admin1@example.com" -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            Add-RecipientPermission -Identity $UserPrincipalName -Trustee "admin2@example.com" -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            Add-RecipientPermission -Identity $UserPrincipalName -Trustee "service@example.onmicrosoft.com" -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            
            # Add 'FullAccess' permissions for a specific user or group
            Add-MailboxPermission -Identity $UserPrincipalName -User "helpdesk@example.com" -AccessRights FullAccess -AutoMapping $false -ErrorAction Stop
            
            # Set the retention policy for the mailbox
            Set-Mailbox -Identity $UserPrincipalName -RetentionPolicy "Default Retention Policy" -ErrorAction Stop
            
            # Disable specific email apps for the user
            Set-CASMailbox -Identity $UserPrincipalName -PopEnabled $false -ImapEnabled $false -ActiveSyncEnabled $false -ErrorAction Stop
            
            Write-Host "Successfully applied mailbox settings for: $UserPrincipalName" -ForegroundColor Green
            
            # If no errors were thrown, break the loop
            break
        }
        catch {
            $TryCount++
            Write-Host "Attempt $TryCount of $MaxTries failed. Retrying..." -ForegroundColor Yellow
            Start-Sleep -Seconds 30  # Wait for 30 seconds before retrying
        }
    }
    while ($TryCount -lt $MaxTries)

    if ($TryCount -eq $MaxTries) {
        Write-Host "All attempts to apply mailbox settings for: $UserPrincipalName have failed." -ForegroundColor Red
        Write-Error $_.Exception.Message
    }
}

Function new-ADUserWithHomeFolder {
    <#
    .DESCRIPTION
    This function creates a home folder for a new AD user and sets the appropriate permissions.

    .PARAMETER Username
    The username of the AD user for whom the home folder is being created.

    .NOTES
    This function assumes you have the necessary permissions to create folders on the file server
    and modify AD user objects. It also assumes the AD module is loaded.
    #>
    param($Username)

    # Define the base path for home folders
    $BaseHomeFolderPath = "\\FileServer\HomeFolders"

    # Construct the full path for the user's home folder
    $HomeFolderPath = Join-Path -Path $BaseHomeFolderPath -ChildPath $Username

    # Set the home directory path for the user in AD
    Set-ADUser $Username -HomeDirectory $HomeFolderPath -HomeDrive "H:"

    try {
        # Create the user's home folder
        New-Item -Path $HomeFolderPath -ItemType Directory -Force -ErrorAction Stop

        # Get the ACL of the new folder
        $Acl = Get-Acl -Path $HomeFolderPath

        # Get the user's SID
        $UserSID = (Get-ADUser $Username).SID

        # Create a new access rule granting the user full control
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $UserSID, 
            "FullControl", 
            "ContainerInherit,ObjectInherit", 
            "None", 
            "Allow"
        )

        # Add the access rule to the ACL
        $Acl.SetAccessRule($AccessRule)

        # Apply the updated ACL to the folder
        Set-Acl -Path $HomeFolderPath -AclObject $Acl

        Write-Host "Home folder created and permissions set for user: $Username" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create home folder or set permissions for user: $Username. Error: $_"
    }
}


Function Add-ADGroups {
    <#
    .DESCRIPTION
    This function adds a user to appropriate AD groups based on their office, title, and department.

    .PARAMETER Username
    The username of the AD user to be added to groups.

    .PARAMETER Office
    The office location of the user.

    .PARAMETER Title
    The job title of the user.

    .PARAMETER Department
    The department of the user.

    .NOTES
    This function assumes you have the necessary permissions to query AD and modify group memberships.
    It also assumes the AD module is loaded.
    #>
    Param(
        [Parameter(Mandatory = $True)]
        [String]$Username, 

        [Parameter(Mandatory = $True)]
        [String]$Office,

        [Parameter(Mandatory = $True)]
        [String]$Title,

        [Parameter(Mandatory = $True)]
        [String]$Department
    )

    # Initialize arrays to store AD users and groups
    $ADUsers = @()
    $ADGroups = @()

    # Search base for AD queries
    $SearchBase = "OU=Users,DC=contoso,DC=com"

    # Query AD for users with matching attributes
    $ADUsers += Get-ADUser -Filter {
        (Department -eq $Department) -and 
        (Title -eq $Title) -and 
        (Office -eq $Office) -and 
        (Enabled -eq $true)
    } -SearchBase $SearchBase -Properties Department, Title, Office, MemberOf, Enabled

    # Also search for users with matching Description instead of Title
    $ADUsers += Get-ADUser -Filter {
        (Department -eq $Department) -and 
        (Description -eq $Title) -and 
        (Office -eq $Office) -and 
        (Enabled -eq $true)
    } -SearchBase $SearchBase -Properties Department, Description, Office, MemberOf, Enabled

    # Identify common groups among matching users
    $GroupCounts = $ADUsers.MemberOf | Group-Object | Where-Object { $_.Count -ge ($ADUsers.Count * 0.7) }
    $ADGroups = $GroupCounts.Name

    # Add the user to each identified group
    foreach ($Group in $ADGroups) {
        try {
            Add-ADGroupMember -Identity $Group -Members $Username -ErrorAction Stop
            Write-Host "Added $Username to group: $Group" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to add $Username to group: $Group. Error: $_"
        }
    }

    if ($ADGroups.Count -eq 0) {
        Write-Warning "No common groups found for user with Office: $Office, Title: $Title, Department: $Department"
    }
}


Function AD-changes {
    <#
    .DESCRIPTION
    This function applies various Active Directory changes for a new user account.

    .PARAMETER NewHire
    An object containing the new hire's information.

    .NOTES
    This function assumes you have the necessary permissions to modify AD objects and that the AD module is loaded.
    #>
    param($NewHire)

    # Define generic groups
    $DefaultGroups = @("All Users", "New Hires")
    $SpecialGroups = @("External Access")

    # Add user to default groups
    foreach ($group in $DefaultGroups) {
        try {
            Add-ADGroupMember -Identity $group -Members $NewHire.Username -ErrorAction Stop
            Write-Host "Added $($NewHire.Username) to group: $group" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to add $($NewHire.Username) to group: $group. Error: $_"
        }
    }

    # Add user to special groups based on status
    if ($NewHire.Status -eq "Contractor") {
        foreach ($group in $SpecialGroups) {
            try {
                Add-ADGroupMember -Identity $group -Members $NewHire.Username -ErrorAction Stop
                Write-Host "Added $($NewHire.Username) to special group: $group" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to add $($NewHire.Username) to special group: $group. Error: $_"
            }
        }
    }

    # Set additional AD attributes
    try {
        $adUserParams = @{
            Identity = $NewHire.Username
            Replace  = @{
                'customAttribute1' = $NewHire.CellPhoneNumber
                'customAttribute2' = $NewHire.Department
                'customAttribute3' = $NewHire.JobTitle
            }
        }
        Set-ADUser @adUserParams -ErrorAction Stop
        Write-Host "Updated custom attributes for $($NewHire.Username)" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to update custom attributes for $($NewHire.Username). Error: $_"
    }

    # Set SMTP address
    $smtpAddress = "SMTP:" + $NewHire.Username + "@contoso.com"
    try {
        Set-ADUser -Identity $NewHire.Username -Add @{ProxyAddresses = $smtpAddress} -ErrorAction Stop
        Write-Host "Set SMTP address for $($NewHire.Username)" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to set SMTP address for $($NewHire.Username). Error: $_"
    }

    # Add user to groups based on attributes
    Add-ADGroups -Username $NewHire.Username -Office $NewHire.Office -Title $NewHire.JobTitle -Department $NewHire.Department

    # Set logon script
    Set-LogonScript -NewHire $NewHire

    # Manage phone number visibility
    Manage-ADUserMobilePhone -UserPrincipalName $NewHire.Username -MobilePhone $NewHire.PersonalNumber -showNumber $NewHire.ShowNumber

    # Add certifications if applicable
    if ($NewHire.Certification) {
        try {
            Set-ADUser -Identity $NewHire.Username -Add @{CustomAttribute4 = $NewHire.Certification} -ErrorAction Stop
            Write-Host "Added certification information for $($NewHire.Username)" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to add certification information for $($NewHire.Username). Error: $_"
        }
    }
}

# ----------------------------------[     Controller     ]-------------------------------------#

Function NewHire {
    <#
    .DESCRIPTION
    This function orchestrates the new hire onboarding process, including retrieving new hire information,
    checking licenses, creating AD accounts, and performing various setup tasks.

    .NOTES
    This function assumes you have the necessary permissions and connections to SharePoint,
    Exchange Online, and other relevant services.
    #>

    # Clear the console for better readability
    Clear-Host

    # Retrieve new hire information from SharePoint
    $NewHireListItems = Get-PnPListItem -List "New Hire Onboarding List"
    $Global:NewHireList = foreach ($Item in $NewHireListItems) {
        [PSCustomObject]@{
            ID                = $Item["ID"]
            FirstName         = $Item["FirstName"]
            LastName          = $Item["LastName"]
            StartDate         = $Item["StartDate"].ToString("MM/dd/yyyy")
            Department        = $Item["Department"]
            JobTitle          = $Item["JobTitle"]
            Status            = $Item["EmploymentStatus"]
            Office            = $Item["Office"]
            ApprovalStatus    = $Item["ApprovalStatus"]
            Expedite          = $Item["Expedite"]
            PersonalNumber    = $Item["PersonalPhoneNumber"]
            ShowNumber        = $Item["DisplayPhoneNumber"]
            EmployeeID        = $Item["EmployeeID"]
            Certification     = $Item["Certification"]
            ReturningEmployee = $Item["ReturningEmployee"]
        }
    }

    # Filter new hires based on start date and approval status
    $Global:NewHireList = $Global:NewHireList | Where-Object {
        $startDate = [datetime]$_.StartDate
        $daysUntilStart = (New-TimeSpan -Start (Get-Date) -End $startDate).Days
        (($daysUntilStart -ge 0 -and $daysUntilStart -lt 21) -or $_.Expedite -eq "Yes") -and
        $_.ApprovalStatus -eq "Approved" -and
        $_.ReturningEmployee -ne "Yes"
    }

    # Check and allocate licenses
    Write-Host "Checking license availability..." -ForegroundColor Cyan
    License-Count
    $Global:NewHireList = $Global:NewHireList | ForEach-Object { License-check -NewHireList $_ }

    # Create AD accounts
    Write-Host "Creating AD accounts..." -ForegroundColor Cyan
    $Global:NewHireList = $Global:NewHireList | ForEach-Object { AD-Account -NewHire $_ }

    # Sync AD accounts to Azure AD and assign licenses
    Write-Host "Syncing to Azure AD and assigning licenses..." -ForegroundColor Cyan
    foreach ($NewHire in $Global:NewHireList) {
        Waitfor-sync -Email $NewHire.Email -status $NewHire.Status -Username $NewHire.Username
    }

    # Wait for mailboxes to be created
    Write-Host "Waiting for mailboxes to be created..." -ForegroundColor Cyan
    foreach ($NewHire in $Global:NewHireList) {
        Waitfor-Mailbox -Email $NewHire.Email -Username $NewHire.Username
    }

    # Perform additional setup tasks for each new hire
    foreach ($NewHire in $Global:NewHireList) {
        Write-Host "Performing additional setup for $($NewHire.FirstName) $($NewHire.LastName)..." -ForegroundColor Cyan

        # Update AD attributes and group memberships
        AD-changes -NewHire $NewHire

        # Configure mailbox settings
        MailboxSettings -UserPrincipalName $NewHire.Email

        # Create home folder (if applicable)
        if ($NewHire.Office -ne "Remote") {
            new-ADUserWithHomeFolder -Username $NewHire.Username
        }

        # Update SharePoint list status
        Set-PnPListItem -List "New Hire Onboarding List" -Identity $NewHire.ID -Values @{
            "Status" = "Provisioned"
            "Email"  = $NewHire.Email
        }

        # Add to onboarding workflow list
        Add-PnPListItem -List "Onboarding Workflow" -Values @{
            'NewHire'    = $NewHire.Email
            'StartDate'  = $NewHire.StartDate
            'Department' = $NewHire.Department
            'Office'     = $NewHire.Office
        }
    }

    # Send notification emails
    Email-Notification
}



#Run the scripts
NewHire
Start-Sleep -Seconds (30)
