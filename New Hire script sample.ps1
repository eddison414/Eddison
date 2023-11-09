
<#
Please be advised that this script has been specifically developed for our unique IT infrastructure. It has been customized to align with the specific requirements and configurations of our system. Consequently, the script contains domain-specific information that has been adapted to suit our environment.

It is essential to understand that due to these customizations, the script is not designed for universal application. Attempting to implement it in a different IT environment may not yield the intended results and could potentially lead to system incompatibilities or operational issues.

We advise against using this script in any IT setting other than the one it was specifically created for. If you have any questions or require further clarification, please do not hesitate to reach out.
#>﻿

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
    #Install-Module -Name PowerShellGet -Force
    #Install-Module -Name Teams -Force
    Install-Module -Name PowerShellGet -Force
    install-Module -Name MicrosoftTeams 
}

Function Encrypt {
    #### Set and encrypt our own password to file using default ConvertFrom-SecureString method
(get-credential).password | ConvertFrom-SecureString | set-content 'C:\Users\Automation\Documents\PowerShell Automation\E.txt'
}
    
# ----------------------------------[ Login into the require enviroment]-------------------------------

$prem = New-Object System.Management.Automation.PsCredential("automation", $encrypted)
$Cloud = New-Object System.Management.Automation.PsCredential("Account", $encrypted)

#Connect
Connect-PnPOnline -Url "https://ocdcrm.sharepoint.com/sites/DGCCPAClassic/IT" -Credentials $Cloud
Initialize-iPilotSession -ApiKey 'auwpy2oQeK5KUYXvJAhid1ekgi8GrfesaC1qn0uy' -Credential $Cloud
Connect-MsolService -Credential $Cloud
Connect-ExchangeOnline -Credential $Cloud
Connect-AzureAD -Credential $Cloud
Connect-MicrosoftTeams -Credential $Cloud

#Test for connection making sure at least 1 of the DCs is online
$Connection = Test-Connection -ComputerName 'HA-DC1'
IF ($Connection) { $PSSessionDC1 = New-PSSession -ComputerName 'HA-DC1' -Credential $prem }
IF (!($Connection)) {
    $Connection = Test-Connection -ComputerName 'PKF-EIP-DC1.odmd.local'
    IF ($Connection) { $PSSessionDC1 = New-PSSession -ComputerName 'HA-DC1' -Credential $prem }
}
IF (!($Connection)) { $PSSessionDC1 = New-PSSession -ComputerName 'NYM-DC1' -Credential $prem }


$ichannel = New-PSSession -ComputerName 'Ichannel' -Credential $prem
    
# ----------------------------------[     Functions     ]-------------------------------------
  
Function AD-Account {
    param($NewHire)

    if ($NewHire.LicenseCheck -eq 'Pass') {

        $DisplayMame = if ($NewHire.Middle) { $NewHire.FirstName + " " + $NewHire.Middle + " " + $NewHire.LastName }else { $NewHire.FirstName + " " + $NewHire.LastName }
        $UserPrincipalName = Username-Check -firstName $NewHire.FirstName -LastName $NewHire.LastName
        $UserPrincipalName = $UserPrincipalName.ToLower()

        # Created the AD account in the DC - DGCDOM1
        $Success = Invoke-Command -Session $PSSessionDC1 -ScriptBlock {

            $Success = $False; $TryCount = 0; $MaxTries = 5
            while (!($Success) -and $TryCount -le $MaxTries) {
                try {            
                    # 	Create a new Active Directory Account 
                    New-ADUser -Name  $args[0] `
                        -DisplayName $args[0] `
                        -SamAccountName $args[1]`
                        -Path $args[2]  `
                        -UserPrincipalName $args[3] `
                        -AccountPassword $args[4]`
                        -EmailAddress $args[3] `
                        -GivenName  $args[5]`
                        -Surname  $args[6]`
                        -Office $args[7] `
                        -Enabled:$True  `
                        -Description $args[8]  `
                        -Title $args[8]`
                        -Department $args[9] `
                        -Company "PKF O'Connor Davies" `
                        -HomePage "www.pkfod.com" -EmployeeID $args[10] 

                    $Success = $True
                }
                catch {
                    # Increase count by 1
                    $TryCount++                    
                    $Error | Out-File '\\pkfod-automate\C$\temp\NewUserErrorLog.txt' -Append
                }
               
            }

            $Success

        } -ArgumentList ($DisplayMame), # 0
                ($UserPrincipalName), # 1
                (Get-OU -NewHire $NewHire), # 2 
                ($UserPrincipalName + "domain"), # 3
                (Create-Password -StartDate $NewHire.StartDate), # 4
        $NewHire.FirstName, # 5
        $NewHire.LastName, # 6
        (AD-Office -Location $NewHire.City), # 7
        $NewHire.JobTitle, # 8  
        $NewHire.Department, # 9
        $NewHire.EmployeeID #10
        

        if ($Success -eq 'True') {
    
            $NewHire | Add-Member -NotePropertyMembers @{'Username' = $UserPrincipalName;
                'Email'                                             = ($UserPrincipalName + '@domain.com')
            }
            $NewHire  
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
    Param($NewHire)
    # Use switch statement to select the logon script based on the department
    switch ($NewHire.department) {
        # If the department is IT, assign IT logon script
        "IT" { $LogonScript = 'ITlogon-65.bat' }
    }
    # Use switch statement to select the logon script based on the title
    switch ($NewHire.title) {
        # If the title matches Human Resources, assign HR logon script
        "Human Resources*" { $LogonScript = 'HRlogon-65.bat' }
    }
    # Use switch statement to select the logon script based on the office
    switch ($NewHire.office) {
        # Assign office specific logon scripts
        "NJ - Woodcliff Lake" { $LogonScript = 'logon_nj-65.bat' } # NJ - Woodcliff Lake office
        "NJ - Cranford" { $LogonScript = 'logon-65-cr.bat' } # NJ - Cranford office or Clear Thinking
        "international" { $LogonScript = 'logon-65-India.bat' } # ID - Mumbai office
        "NY - Hauppauge" { $LogonScript = 'logon-65-HAP.bat' } # NY - Hauppauge office
        "NY - Newburgh" { $LogonScript = 'logon-65-NB.bat' } # NY - Newburgh office
        "CT - Shelton" { $LogonScript = 'logon-SHL.bat' } # CT - Shelton office
        "NJ - LBG" { $LogonScript = 'logon_LBG-65.bat' } # NJ - LBG office
        "RI - Providence" { $LogonScript = 'logon-65-BFMM.bat' } # RI - Providence office
        "NY - Middletown" { $LogonScript = 'logon_jgs-65.bat' } # NY - Middletown office
        "MD - Bethesda" { $LogonScript = 'logon-65-md.bat' } # MD - Bethesda office
        "CT - Wethersfield" { $LogonScript = 'logon_wo-65.bat' } # CT - Wethersfield office

        # NJ - Woodcliff Lake
        # those twp logon scripts are assign to the same office. We need to know what is the difference
        "NJ - Woodcliff Lake" { $LogonScript = 'logon_nj-65.bat' }

        # NY - NYC
        # those twp logon scripts are assign to the same office. We need to know what is the difference
        "NY - NYC" { $LogonScript = 'TAX-65.bat' }

        # Assign default logon scripts where there is no match
        "" { $LogonScript = 'logon-65-mktg.bat' } # No office provided, assign marketing logon script
        "" { $LogonScript = 'NTUsers-65.bat' } # No office provided, assign general user logon script

        # Assign default logon script for any other office
        default { $LogonScript = 'logon-65.bat' }
    }
    # Set the script path for the new hire
    Set-ADUser -Identity $NewHire.Username -ScriptPath $LogonScript
}
  
Function Get-OU {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true)]
        [AllowEmptyString ()]
        $NewHire
        
    )
    switch ($NewHire.City) {
        "New York" { $OU = 'OU=NYC,OU=ODMD,DC=odmd,DC=local' ; }
        "Shelton" { $OU = "OU=Shelton,OU=ODMD,DC=odmd,DC=local" ; }
        "Stamford" { $OU = "OU=Stamford,OU=ODMD,DC=odmd,DC=local" ; }
        "Cranford" { $OU = "OU=Cranford,OU=ODMD,DC=odmd,DC=local" ; }
        "Bethesda" { $OU = "OU=Bethesda,OU=ODMD,DC=odmd,DC=local" ; }
        "Hauppauge" { $OU = "OU=Hauppauge,OU=ODMD,DC=odmd,DC=local" ; }
        "Boston" { $OU = "OU=Boston,OU=DGC,OU=ODMD,DC=odmd,DC=local" ; }
        "Woburn" { $OU = "OU=Woburn,OU=DGC,OU=ODMD,DC=odmd,DC=local" ; }
        "Middletown" { $OU = "OU=Middletown,OU=ODMD,DC=odmd,DC=local" ; }
        "Providence" { $OU = "OU=Providence,OU=ODMD,DC=odmd,DC=local" ; }
        "Poughkeepsie" { $OU = "OU=Poughkeepsie,OU=ODMD,DC=odmd,DC=local" ; }
        "Harrison" { $OU = "OU=Harrison,OU=ODMD,DC=odmd,DC=local" ; }
        "Wethersfield" { $OU = "OU=Wethersfield,OU=ODMD,DC=odmd,DC=local" ; }
        "Newburgh" { $OU = "OU=Newburgh,OU=ODMD,DC=odmd,DC=local" ; }
        "Mumbai" { $OU = "OU=OPSEU,OU=Opseu-Outside users,OU=OPSEU-Family Office,OU=ODMD,DC=odmd,DC=local" ; }
        default { $OU = 'OU=NYC,OU=ODMD,DC=odmd,DC=local' }
    }

    # Exections, in case the user should be on a different OU regarless of the location
    <# switch ($NewHire.Department) {
        "IT" { $OU = "OU=Group IT,OU=ODMD,DC=odmd,DC=local" ; break }
        "finance" { $OU = "OU=Financial Services,OU=ODMD,DC=odmd,DC=local" }
        "Family Office" { $OU = "OU=Family Office,OU=ODMD,DC=odmd,DC=local" }
        "Elite Accounting" { $OU = "OU=Emmaus,OU=ODMD,DC=odmd,DC=local" }
        "Clear thinking" { $OU = "OU=Clear Thinking,OU=ODMD,DC=odmd,DC=local" }
    }
    #>

    $OU
}

function AD-Office {
    param(

        [Parameter(Mandatory = $True)]
        [AllowEmptyString ()]
        [String]$Location

    )

    switch ($Location) {
        # New York State
        { "Newburgh", "Middletown", "Poughkeepsie", "Harrison" -eq $_ } { $Office = "NY - $Location" }

        # New Jersey State
        { "Cranford", "Hauppauge", "Woodcliff Lake" -eq $_ } { $Office = "NJ - $Location" }

        # Massachusetts State
        { "Boston", "Woburn" -eq $_ } { $Office = "MA - $Location" }

        # Connecticut State
        { "Shelton", "Stamford", "Wethersfield" -eq $_ } { $Office = "CT - $Location" }

        # Maryland State
        { "Bethesda" -eq $_ } { $Office = "MD - $Location" }

        # Rhode Island
        { "Providence" -eq $_ } { $Office = "RI - $Location" }

        #International
        { "South Africa", "Mumbai", "Philippines" -eq $_ } { $Office = "international" }

        # Default
        default { $Office = "NY - NYC" }
    }

    # Output 
    $Office
}

Function Waitfor-sycn {
    [CmdletBinding()]
    <#
 The script "Waitfor-sycn" is a function that takes in three parameters: an email,
 a username, and a status. It attempts to add the user's email to an Azure AD group based on the status, 
 and retries up to 10 times with a 5-minute sleep in between each failure. If it fails 10 times, 
 it removes the AD account and sends an email to IT. The function also includes a verbose message that logs the progress of the script.
#>
    
    Param([string]$Email, $Username, $status)



    $Success = $False
    $TryCount = 0
    $MaxTries = 10
    $Waiting = 300
    while (!($Success) -and $TryCount -le $MaxTries) {
        try {
            # Assign the user to 'MFA Exclusion Group'
            Add-AzureADGroupMember -ObjectId '28ade99f-e283-4442-96ec-2437279e8d5f' -RefObjectId (Get-AzureADUser -ObjectId $Email).ObjectId -Verbose
            
            if ($status -like "Employee (Full or Part Time)") {
                #  Assign Microsft 365 E3, Teams phone Standard, and conferencing
                Add-AzureADGroupMember -ObjectId '8ccc4f3f-beb7-45b2-9c79-4b5174ceef3b' -RefObjectId (Get-AzureADUser -ObjectId $Email).ObjectId -Verbose
                Write-Color -Text "The user: ", "$Username ", "has been added to the Azure group: ", "Standard US User" -Color White, Green, White, Green
                # Write-Verbose -Message "The user $Email, has been added to the group: Microsoft 365 E3 License"
            }
            else {
                #  Assign Microsft 365 E3, and conferencing
                Add-AzureADGroupMember -ObjectId 'c29f8193-9b6f-4a40-935b-1e5c8697e8b5' -RefObjectId (Get-AzureADUser -ObjectId $Email).ObjectId -Verbose
                Write-Color -Text "The user: ", "$Username ", "has been added to the Azure group:", "Interns, US Contractors, and Offshore Users" -Color White, Green, White, Green
                #Write-Verbose -Message "The user $Email, has been added to the group: Microsoft 365 E3 License"
            
            }
    
            $Success = $True
        }
        catch {


            if ($TryCount -eq 2) { $Null = Invoke-Command -Session $PSSessionDC1 -ScriptBlock { start-adsyncsynccycle } }
            if ($TryCount -eq 5) { $Null = Invoke-Command -Session $PSSessionDC1 -ScriptBlock { start-adsyncsynccycle } }
            if ($TryCount -eq 8) { $Null = Invoke-Command -Session $PSSessionDC1 -ScriptBlock { start-adsyncsynccycle } }
            if ($TryCount -le 9) {
                # If it fail
                Write-Color -text "The user ", "$Email ", "has not been found in the cloud, waiting ", "$($Waiting / 60) ", "minutes" -Color White, Green, White, Green, White 
                $Waiting = ($Waiting + 300)
                           
                Start-Sleep -Seconds ($Waiting)
            }

            # If the count was to 10 killswitch  the AD Account
            if ($TryCount -eq 10) {
                Remove-ADUser -Server 'HA-DC1' -Identity $Username -confirm:$false; Clear-Variable $NewHire
                Write-Color -Text "The script has done ", "$TryCount ", "tries to find the new hire: ", "$Username", " in the cloud.", "The AD Account will be deleted", " And the script will try again tomorrow" `
                    -Color White, Yellow, White, Green, White, Red, White
                
                # create a script to send an email letting IT know about this.
                $Global:killswitchMessage += "
                A critical step has fail after multiple attempts. The AD account $Username will be deleted, And the script will try again tomorrow. Function; Waitfor-sycn 
                "
            }

            $TryCount++

        }
    
        # Increase count by 1
        # $TryCount++ look into this
    }
}

Function License-Count {
    # Write-Verbose -Message "Geting license count for 'Office E3'"
    $OfficeE3 = Get-MsolAccountSku | where { $_.AccountSkuId -like "OCDCrm:SPE_E3" }
    $Global:OfficeE3 = $OfficeE3.ActiveUnits - $OfficeE3.ConsumedUnits;
    Write-Color -Text 'Geting total free license count for; ', 'Office E3', ' (', "$Global:OfficeE3", ')' -Color White, Yellow, White, Green, White
    
    # Write-Verbose -Message "Geting license count for 'Microsft Teams Phone Standard'"
    $TeamsPhone = Get-MsolAccountSku | where { $_.AccountSkuId -like "OCDCrm:MCOEV" }
    $Global:TeamsPhone = $TeamsPhone.ActiveUnits - $TeamsPhone.ConsumedUnits;
    Write-Color -Text 'Geting total free license count for; ', 'Microsft Teams Phone Standard', ' (', "$Global:TeamsPhone", ')' -Color White, Yellow, White, Green, White
    
    # Write-Verbose -Message "Geting license count for 'Microsoft Teams Audio Conferencing select dial-out'"
    $Conferencing = Get-MsolAccountSku | where { $_.AccountSkuId -like "OCDCrm:Microsoft_Teams_Audio_Conferencing_select_dial_out" }
    $Global:Conferencing = $Conferencing.ActiveUnits - $Conferencing.ConsumedUnits;
    Write-Color -Text 'Geting total free license count for; ', 'Microsoft Teams Audio Conferencing select dial-out', ' (', "$Global:Conferencing", ')' -Color White, Yellow, White, Green, White



    $Global:Request_OfficeE3 = 0
    $Global:Request_TeamsPhone = 0
    $Global:Request_Conferencing = 0

}
         
function License-check {
    param($NewHireList)

                            
    Begin {  }

    Process {
            
        foreach ($User in $NewHireList) {
            $FirstName = $user.FirstName
            
            $Global:OfficeE3 = ($Global:OfficeE3 - 1);
            if ($Global:OfficeE3 -gt 1) {
                Write-Color -Text 'Office E3;', ' (', "$Global:OfficeE3", ') ', 'License check for the user; ', "$($user.FirstName + ' ' + $user.lastName)", "...", "PASS!" `
                    -Color yellow, White, Green, White, White, Yellow, White, Green
                $E3 = 'True'
            }
            else {
                Write-Color -Text 'Office E3;', ' (', "$Global:OfficeE3", ') ', 'License check for the user; ', "$($user.FirstName + ' ' + $user.lastName)", "...", "FAIL!" `
                    -Color yellow, White, Red, White, White, Yellow, White, Red
                $Global:Request_OfficeE3 += 1
            }

            if ($User.Status -eq "Employee (Full or Part Time)") {
                $Global:TeamsPhone = ($Global:TeamsPhone - 1);
                if ($Global:TeamsPhone -gt 1) {
                    Write-Color -Text 'Teams Phone Standard;', ' (', "$Global:TeamsPhone", ') ', 'License check for the user; ', "$($user.FirstName + ' ' + $user.lastName)", "...", "PASS!" `
                        -Color yellow, White, Green, White, White, Yellow, White, Green
                    $Teams = 'True'
                }
                else {
                    Write-Color -Text 'Teams Phone Standard;', ' (', "$Global:TeamsPhone", ') ', 'License check for the user; ', "$($user.FirstName + ' ' + $user.lastName)", "...", "FAIL!" `
                        -Color yellow, White, Red, White, White, Yellow, White, Red
                    $Global:Request_TeamsPhone += 1
                } 
            }

            
            $Global:Conferencing = ($Global:Conferencing - 1);
            if ($Global:Conferencing -gt 1) {
                Write-Color -Text 'Microsoft Teams Audio Conferencing select dial-out;', ' (', "$Global:Conferencing", ') ', 'License check for the user; ', "$($user.FirstName + ' ' + $user.lastName)", "...", "PASS!" `
                    -Color yellow, White, Green, White, White, Yellow, White, Green
                $Conference = 'True'
            }
            else {
                Write-Color -Text 'Microsoft Teams Audio Conferencing select dial-out;', ' (', "$Global:Conferencing", ') ', 'License check for the user; ', "$($user.FirstName + ' ' + $user.lastName)", "...", "FAIL!" `
                    -Color yellow, White, Red, White, White, Yellow, White, Red
                $Global:Request_Conferencing += 1
            }     
            
            if (($User.Status -eq "Employee (Full or Part Time)") -and (($E3 -eq 'True') -and ($Teams -eq 'True') -and ($Conference -eq 'True'))) { 
                Write-Color "the user, ", "$($user.FirstName + ' ' + $user.lastName) ", "has pass ", "(3/3) ", "License check"  -Color White, Green, White, Green, White
                $User  | Add-Member -NotePropertyMembers @{'LicenseCheck' = "Pass"; }
                $User  
            } 
            if ((($User.Status -eq 'Contractor/Consultant') -or ($User.Status -eq 'Intern')) -and (($E3 -eq 'True') -and ($Conference -eq 'True'))) {
                Write-Color "the user, ", "$($user.FirstName + ' ' + $user.lastName) ", "has pass ", "(2/2) ", "License check"  `
                    -Color White, Green, White, Green, White
                $User  | Add-Member -NotePropertyMembers @{'LicenseCheck' = "Pass"; }
                $User             
            }
            else {
                Write-Color -Text "The user; ", "$($user.FirstName + ' ' + $user.lastName)", "becasuse there are not suficient licenses for it" `
                    -Color White, Yellow, White
            }
        }
    
  
    
        
    }

    End { 

             
    }
}

Function Waitfor-Mailbox {
    [CmdletBinding()]
    Param([string]$Email, $Username)

    $Success = $False; $TryCount = 0; $MaxTries = 10
    while (!($Success) -and $TryCount -le $MaxTries) {
        try {
            #  Script goes here
            $Mailbox = Get-Mailbox -Identity $Email -ErrorAction SilentlyContinue 
            
            if ($Mailbox) { 
                Write-Host "Mailbox for $Email has been created"
                $Success = "True" 
            }  
           
        }
        catch {
            # If the count was to 10 killswitch  the AD Account
            if ($TryCount -eq 10) {
                Remove-ADUser -Server 'HA-DC1' -Identity $Username -confirm:$false; Clear-Variable $NewHire
                Write-Color -Text "The script has done ", "$TryCount ", "tries to find the new hire: ", "$Username", " in the cloud.", "The AD Account will be deleted", " And the script will try again tomorrow" `
                    -Color White, Yellow, White, Green, White, Red, White
                
                # create a script to send an email letting IT know about this.
                $Global:killswitchMessage += "
                A critical step has fail after multiple attempts. The AD account $Username will be deleted, And the script will try again tomorrow. Function; Waitfor-Mailbox 
                "
            }
            # If it fail
            Write-Host "The mailbox for  $Email, has not been created, waiting 2 minute"
            Start-Sleep -Seconds (60) 
            $TryCount++
        }
    }
}

function Username-Check {
    <#
    This is a PowerShell function named "Username-Check" that generates a unique email by 
    concatenating the first character of the first name, the last name, and a domain name.
    It checks if the email is already taken in Active Directory and, if so, 
    tries again with a dot between the first name and last name. If both attempts fail, 
    it returns a warning message. The function takes two mandatory string parameters: "firstName" and "lastName".
    #>
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [String]$firstName,

        # Param2 help description
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [String]$lastName
    )

    $domain = "@domain.com"
    $email = $firstName[0] + $lastName + $domain

    # Check for unique email in AD
    $Email1 = Get-ADObject -Filter { mail -eq $email }

    if ($Email1) {
        $email = $firstName + '.' + $lastName + $domain
        $Email2 = Get-ADObject -Filter { mail -eq $email }

        if ($Email2) {
            Write-Warning "Cannot generate a unique email for $firstName $lastName"
        }
        else {
            $email = $email
        }
    }
    else {
        $email = $email
    }

    $username, $domain = $email -split '@'

    $username
}

Function Create-Password {
    param(
        $StartDate
    )

    function Get-DateSuffix([datetime]$Date) {
        switch -regex ($Date.Day.ToString()) {
            '1(1|2|3)$' { 'th' }
            '.?1$' { 'st' }
            '.?2$' { 'nd' }
            '.?3$' { 'rd' }
            default { 'th' }
        }
    }
    [datetime]$StartDate = $StartDate 
    [String]$password = "{0:MMM}{1}{2}{3}" -f $StartDate, $StartDate.Day, (Get-DateSuffix $StartDate), $StartDate.Year
    [String]$Pass = @('domain' + $password + '!!')

    $PW = $Pass | ConvertTo-SecureString -AsPlainText -Force
    $PW 
}

function change-physicaladdress {
    <#
This script is a PowerShell function called "change-physicaladdress" that takes a single parameter, 
a user's username, and updates the user's physical address in Active Directory based on their office location.

The function first uses the Get-ADUser cmdlet to retrieve the user's office location from Active Directory. 
It then uses a switch statement to determine which office location the user is in and sets the corresponding street address, city, state, zip code, and front desk phone number.

Each case in the switch statement corresponds to a different office location. For example, 
if the user's office location is "NY - NYC", the street address is set to "245 Park Avenue, 12th Floor", 
the city is set to "New York", the state is set to "NY", the zip code is set to "10167", and the front desk phone number is set to "212.286.2600".

Once the variables are set, the function then uses the Set-ADUser cmdlet to update the user's physical address in Active Directory with the values stored in the variables.
It is important to note that this script is not complete and does not include the actual code to update the user's physical address in Active Directory, 
and also the switch statement does not cover all states.
#>

    param(

        [Parameter(Mandatory = $True)]
        [AllowEmptyString ()]
        [String]$Username
    )
    
    $Location = (Get-ADUser -Identity $Username -Properties office -Server HA-DC1).office

    switch ($Location) {
        # New York State
        "NY - NYC" {
            # NY - NYC
            $StreetAddress = '245 Park Avenue, 12th Floor'     
            $City = 'New York' 
            $State = 'NY' 
            $Zip = '10167' 
            $FromtDesk = '212.286.2600'
            ; break
        }
        "NY - Newburgh" {
            # NY - Newburgh
            $StreetAddress = '32 Fostertown Rd'     
            $City = 'Newburgh'
            $State = 'NY'
            $Zip = '12550'
            $FromtDesk = '845.565.5400'
            ; break
        }
        "NY - Middletown" {
            # NY - Middletown
            $StreetAddress = '633 Route 211 East'      
            $City = 'Middletown'
            $State = 'NY'
            $Zip = '10941'
            $FromtDesk = '845.565.5400'
            ; break
        }
        "NY - Poughkeepsie" {
            # NY - Poughkeepsie
            $StreetAddress = '2645 South Road, Suite 5'      
            $City = 'Poughkeepsie'
            $State = 'NY'
            $Zip = '12601'
            $FromtDesk = '845.692.9500'
            ; break
        }
        "NY - Harrison" {
            # NY - Harrison
            $StreetAddress = '500 Mamaroneck Avenue, Suite 301'     
            $City = 'Harrison'
            $State = 'NY'
            $Zip = '10528'
            $FromtDesk = '914.381.8900'
            ; break
        }

        # New Jersey State
        "NJ - Cranford" {
            # NJ - Cranford
            $StreetAddress = '20 Commerce Drive, Suite 301'     
            $City = 'Cranford'
            $State = 'NJ'
            $Zip = '07016'
            $FromtDesk = '908.272.6200'
            ; break
        }
        "NJ - Hackensack" {
            # NJ - Hackensack
            $StreetAddress = '878 Veterans Memorial Highway, 4th Floor'      
            $City = 'Hackensack'
            $State = 'NJ'
            $Zip = '07601'
            $FromtDesk 
            ; break
        }
        "NJ - Woodcliff Lake" {
            # NJ - Woodcliff Lake
            $StreetAddress = '300 Tice Boulevard, Suite 315'      
            $City = 'Woodcliff Lake'
            $State = 'NJ'
            $Zip = '07677'
            $FromtDesk = '201.712.9800'
            ; break       
        
        }

        # Massachusetts State
        "MA - Boston" {
            # MA - Boston
            $StreetAddress = '155 Federal Street, Suite 200'      
            $City = 'Boston'
            $State = 'MA'
            $Zip = '02110'
            $FromtDesk = '781.937.5300'
            ; break
        }
        "MA - Woburn" {
            # MA - Woburn
            $StreetAddress = '150 Presidential Way, Suite 510'      
            $City = 'Woburn'
            $State = 'MA'
            $Zip = '01801'
            $FromtDesk = '781.937.5300'
            ; break
        }

        # Connecticut State
        "CT - Shelton" {
            # CT - Shelton
            $StreetAddress = 'One Corporate Drive, Suite 725'        
            $City = 'Shelton'
            $State = 'CT'
            $Zip = '06484'
            $FromtDesk = '203.929.3535'
            ; break
        }
        "CT - Stamford" {
            # CT - Stamford
            $StreetAddress = '3001 Summer Street - 5th Floor East'      
            $City = 'Stamford'
            $State = 'CT'
            $Zip = '06905'
            $FromtDesk = '203.323.2400'
            ; break
        }
        "CT - Wethersfield" {
            # CT - Wethersfield
            $StreetAddress = '100 Great Meadow Road, Suite 207'     
            $City = 'Wethersfield'
            $State = 'CT'
            $Zip = '06109'
            $FromtDesk = '860.257.1870'
            ; break
        }

        # Maryland State
        "MD - Bethesda" {
            # MD - Bethesda
            $StreetAddress = '2 Bethesda Metro Center, Suite 420'      
            $City = 'Bethesda'
            $State = 'MD'
            $Zip = '20814'
            $FromtDesk = '301.652.3464'
            ; break
        }

        # Rhode Island
        "RI - Providence" {
            # RI - Providence
            $StreetAddress = '40 Westminster Street, Suite 600'      
            $City = 'Providence'
            $State = 'RI'
            $Zip = '02903'
            $FromtDesk = '401.621.6200'
            ; break
        }

        default {
            # NY - NYC
            $StreetAddress = '245 Park Avenue, 12th Floor'     
            $City = 'New York' 
            $State = 'NY' 
            $Zip = '10167' 
            $FromtDesk = '212.286.2600'
        }
    }


    # Update AD Attribute for physicaladdress
    Set-ADUser -Identity $Username `
        -StreetAddress $StreetAddress `
        -City $City `
        -State $State `
        -PostalCode $Zip `
        -Server 'HA-DC1' `
        -Replace @{'extensionAttribute4' = $FromtDesk }
}

Function MailboxSettings {
    <#
    .SYNOPSIS
    This function applies several settings and permissions to a specified mailbox.
    .PARAMETER UserPrincipalName
    The user principal name (UPN) of the user.
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
            $Null = Add-RecipientPermission -Identity $UserPrincipalName -Trustee "email@domain.com" -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
            $Null = Add-RecipientPermission -Identity $UserPrincipalName -Trustee "emailpkfod.com" -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
            $Null = Add-RecipientPermission -Identity $UserPrincipalName -Trustee "email@OCDCrm.onmicrosoft.com" -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
            
            # Add 'FullAccess' permissions for the Prostaff_ExchUser
            $Null = Add-MailboxPermission -Identity $UserPrincipalName -User email@domain -AccessRights FullAccess -AutoMapping $false -ErrorAction SilentlyContinue
            
            # Set the retention policy for the mailbox
            $Null = Set-Mailbox -Identity $UserPrincipalName -RetentionPolicy "PKFOD 18 Month Retention Policy" -ErrorAction SilentlyContinue
            
            # Disable email apps for the user
            $Null = Set-CASMailbox -Identity $UserPrincipalName -PopEnabled $false -ImapEnabled $false -ActiveSyncEnabled $false -ErrorAction SilentlyContinue
            
            Write-Host "Successfully applied mailbox permissions for: $UserPrincipalName" -ForegroundColor Green
            
            # If no errors were thrown, the script will reach this point and break the loop
            break
        }
        catch {
            # If an error was thrown, increment the try count and try again
            $TryCount++
            Write-Host "Attempt $TryCount of $MaxTries failed. Retrying..." -ForegroundColor Yellow
        }
    }
    while ($TryCount -lt $MaxTries)

    if ($TryCount -eq $MaxTries) {
        Write-Host "All attempts to apply mailbox permissions for: $UserPrincipalName have failed." -ForegroundColor Red
    }
}

Function Assign-PhoneNumber {
    param(
        $FirstName,
        $LastName,
        $Email,
        $Username
    )

    Begin {
        $Success = $False
        $TryCount = 0
        $MaxTries = 10

        while (!($Success) -and $TryCount -le $MaxTries) {
            try {
                $TelephoneNumber = Get-iPilotNumber -Available -iPilotDomain nuwms00076 -Credential $Cloud |
                Select-Object UserPrincipalName, firstname, lastname, TelephoneNumber |
                Where-Object { $_.FirstName -eq $null } | Get-Random

                $Success = $True
            }
            catch {
                if ($TryCount -eq 5) {
                    Initialize-iPilotSession -ApiKey 'auwpy2oQeK5KUYXvJAhid1ekgi8GrfesaC1qn0uy' -Credential $Cloud
                }

                $TryCount++
            }
        }
    }

    Process {
        $Success = $False
        $TryCount = 0

        while (!($Success) -and $TryCount -le $MaxTries) {
            try {
                New-iPilotTeamsUserAssignment -UserPrincipalName $Email `
                    -FirstName $FirstName `
                    -LastName $LastName `
                    -telephonenumber $TelephoneNumber.TelephoneNumber `
                    -iPilotDomain nuwms00076 -Credential $Cloud

                $P = $TelephoneNumber.TelephoneNumber
                [String]$P = '+1 ' + $P.Substring(0, 3) + '.' + $P.Substring(3, 3) + '.' + $P.Substring(6)

                $NewHire | Add-Member -NotePropertyMembers @{'PhoneNumber' = $P }

                $Success = $True
            }
            catch {
                if ($TryCount -eq 5) {
                    Initialize-iPilotSession -ApiKey 'auwpy2oQeK5KUYXvJAhid1ekgi8GrfesaC1qn0uy' -Credential $Cloud
                }

                if ($TryCount -eq 10) {
                    Remove-ADUser -Server 'HA-DC1' -Identity $Username -Confirm:$false
                    Clear-Variable $NewHire

                    Write-Color -Text "The script has done ", "$TryCount ", "tries to find the new hire: ", "$Username", " in the cloud.", "The AD Account will be deleted", " And the script will try again tomorrow" `
                        -Color White, Yellow, White, Green, White, Red, White

                    $Global:killswitchMessage += @"
A critical step has failed after multiple attempts. The AD account $Username will be deleted, and the script will try again tomorrow. Function: Assign-PhoneNumber
"@
                }

                $TryCount++
            }
        }
    }

    End {}
}

Function Email-Notification {

    if ($Global:Request_OfficeE3 -gt 1) {
        Write-Verbose -Message "Requesting OfficeE3 licenses"
        $Total = ($Global:Request_OfficeE3 + (Get-MsolAccountSku | where { $_.AccountSkuId -like "OCDCrm:SPE_E3" }).ActiveUnits)
        $Message = "
            Good day, 
            
            We require $Request_OfficeE3 Microsoft Office E3licenses, for a total of $total licenses
    
    
            Thanks
             "
               
        Send-MailMessage -From 'automation@ocdcrm.onmicrosoft.com' `
            -To 'email@domain' `
            -Subject 'Office365 license need to be update' `
            -Credential $Cloud `
            -SmtpServer smtp.office365.com `
            -Port 587 -UseSsl `
            -Body $Message
    }
    
    
    if ($Global:Request_TeamsPhone -gt 1 ) {
        Write-Verbose -Message "Requesting OfficeE3 licenses"
        $Total = ($Global:Request_TeamsPhone + (Get-MsolAccountSku | where { $_.AccountSkuId -like "OCDCrm:MCOEV" }).ActiveUnits)
        $Message = "
            Good day, 
            
            We require $Request_TeamsPhone Microsoft  Microsft Teams Phone Standard licenses, for a total of $total licenses
    
    
            Thanks

            Dear IT Team,

            The new hire onboarding script requires additional Microsoft Teams Phone Standard licenses to complete provisioning for new employees.

            Today, the script processed new hires who need to be assigned Teams Phone licenses based on their job roles.

            Unfortunately,To continue automatically assigning licenses to new hires, I am requesting an additional $Request_TeamsPhone Microsoft Teams Phone Standard licenses, bringing the total to $total.

            Please let me know if you need any extra details or have any concerns with fulfilling this request. I'm happy to provide reports on current license usage and assignment.

            Thank you for your assistance in ensuring we have sufficient licenses available to smoothly onboard new employees. Let me know if you have any other questions!

            Regards

             "
    
        Send-MailMessage -From 'automation@ocdcrm.onmicrosoft.com' `
            -To 'email@domain' `
            -Subject 'Request for Additional Microsoft Teams Phone Licenses' `
            -Credential $Cloud `
            -SmtpServer smtp.office365.com `
            -Port 587 -UseSsl `
            -Body $Message -Priority High
    }
    
     
    if ($Global:Request_Conferencing -gt 1 ) {
        Write-Verbose -Message "Requesting OfficeE3 licenses"
        $Total = ($Global:Request_Conferencing + (Get-MsolAccountSku | where { $_.AccountSkuId -like "OCDCrm:Microsoft_Teams_Audio_Conferencing_select_dial_out" }).ActiveUnits)
        $Message = "
            Good day, 
            
            We require $Request_Conferencing Microsoft Office E3licenses, for a total of $total licenses
    
    
            Thanks
             "
        Send-MailMessage -From 'automation@ocdcrm.onmicrosoft.com' `
            -To 'email@domain' `
            -Subject 'Office365 license need to be update' `
            -Credential $Cloud `
            -SmtpServer smtp.office365.com `
            -Port 587 -UseSsl `
            -Body $Message
    }   

    # Semd email is the user was not able to be created.
    if ($Global:killswitchMessage) {
        Send-MailMessage -From 'automation@ocdcrm.onmicrosoft.com' `
            -To 'email@domain' `
            -Subject 'One of the step fail for the followin user' `
            -Credential $Cloud `
            -SmtpServer smtp.office365.com `
            -Port 587 -UseSsl `
            -Body $Global:killswitchMessage
    }


    $html = $Global:NewHireList | select FirstName, LastName, Email, JobTitle, Department, PhoneNumber, State, City | ConvertTo-HtmL | Out-String
    $newhtml = $html #-replace "<td>OFFLINE</td>","<td bgcolor=#FF0000'>OFFLINE</td>" -replace "<td>ONLINE</td>","<td bgcolor=#00FF00'>ONLINE</td>"
    Send-MailMessage -From 'automation@ocdcrm.onmicrosoft.com' `
        -To 'NewHireLog@pkfod.com' `
        -Subject 'The following newhires has been created by the newhire script' `
        -Credential $Cloud `
        -SmtpServer smtp.office365.com `
        -Port 587 -UseSsl `
        -Body $newhtml -BodyAsHtml

}

Function Ichannel-Account {
    <#
This function will create a Ichannel 'Account' for internal users by calling a SQL Query from the Icahnnel server

#>
    param(
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [String]$cUserID, # Ichannel login user ID
    
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [String]$cLastName, # user's lastNmae
  
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [String]$cFirstName, # user's FirstName
  
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [String]$cEmail # user's Email
    )
  
    # Create Ichannel account
    Write-Color -Text "Creating Ichannel account for the user: ", "$cEmail" -Color White, Yellow
    Invoke-Sqlcmd -Query "exec cadoc_system.dbo.spCreateInternalSubscriber 'ROOT', '$cUserID','@cPassword','$cLastName','$cFirstName','$cEmail'" `
        -As DataSet -ServerInstance "pkf-eip-sql1.odmd.local\ICHANNEL" -Database "cadoc_system"
    
}

Function CreateApiContext {
    
    # Load required assemblies
    $Null = [System.Reflection.Assembly]::LoadFrom("C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Pfx.Engagement.API\v4.0_2022.1.1.1__21b98a3ae763e7ad\Pfx.Engagement.API.dll")

    $secureString = Get-Content 'C:\Windows\E.txt' | ConvertTo-SecureString
    $E = (New-Object System.Management.Automation.PsCredential("E", $secureString)).Password
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($E)
    $E = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($ptr)
    # Remember to free the BSTR to avoid a memory leak
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    $IsCentral = $true;
    $CreateApiContext = [Pfx.Engagement.API.ApiFactory]::CreateApiContext("HA-SQL4", $IsCentral, "ADMIN", "$E")
    $CreateApiContext
}

Function New-EngagementUser {
    Param([Parameter(Mandatory = $True)]
        $newHire,
        [Parameter(Mandatory = $True)]
        $CreateApiContext
    )

    # Load required assemblies
    $Null = [System.Reflection.Assembly]::LoadFrom("C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Pfx.Engagement.API\v4.0_2022.1.1.1__21b98a3ae763e7ad\Pfx.Engagement.API.dll")

    # Retrieves all the installed license guids with their information in the given api context
    $ILicenseApi = [Pfx.Engagement.API.ApiFactory]::CreateLicenseApi()
    $EngagementLicenseCount = $ILicenseApi.GetAllLicenseTypes($CreateApiContext)

    # Creates an instance of IStaffApi
    $CreateStaffApi = [Pfx.Engagement.API.ApiFactory]::CreateStaffApi()

    # Make sure there are suficient 'Workpaper Management' and 'Trial Balance' licenses to create the user
    $TrialBalance = $EngagementLicenseCount | Where { $_.ProductName -like "Engagement Trial Balance" }
    $WorkpaperManagement = $EngagementLicenseCount | Where { $_.ProductName -like "Engagement Workpaper Management" }
    if (($TrialBalance.AvailableLicenses -gt 0) -and ($WorkpaperManagement.AvailableLicenses -gt 0)) {
        
        switch -regex ($NewHire.Department) {
            "^Admin" { $DepartmentId = 1 }
            "^Audit" { $DepartmentId = 2 }
            "^Tax" { $DepartmentId = 3 }
            "^Advisory" { $DepartmentId = 4 }
            default { $DepartmentId = 0 }
        }


        switch -regex ($NewHire.JobTitle) {
            "^Associate" { $RightsGroupId = 64 }
            "^Senior" { $RightsGroupId = 65 }
            "^Supervisor" { $RightsGroupId = 66 }
            "^Manager" { $RightsGroupId = 66 }
            "^Partner" { $RightsGroupId = 67 }
            default { $RightsGroupId = 64 }
        }

        # Get new hire properties
        $newStaffRequest = New-Object -TypeName Pfx.Engagement.API.Staff.StaffAddRequestDto -Property @{
            RightsGroupId = $RightsGroupId
            Active        = $true
            FirstName     = $newHire.FirstName
            LastName      = $newHire.LastName
            StaffInitial  = ($newHire.FirstName[0] + $newHire.FirstName[1] + "." + $newHire.LastName[0] + $newHire.LastName[1])
            WorkeMail     = $newHire.Email
            MiddleName    = ""
            Login         = $newHire.Username
            DepartmentId  = $DepartmentId
            StaffTitleId  = ""
            PersonalTitle = ""
            HomeEmail     = $newHire.Email
            PhoneNumber   = $newHire.PersonalNumber
            MachineName   = "HA-SQL4"
        }
    
        # Create a new Engagement User
        $staffApi = [Pfx.Engagement.API.ApiFactory]::CreateStaffApi()
        $staffApi.Add($CreateApiContext, $newStaffRequest)

        # Get the user StaffGuid
        $GetAll = $CreateStaffApi.GetAll($CreateApiContext)
        $NewStaff = $GetAll  | where { $_.WorkEmail -like "$($newHire.Email)" }

        # Assign the user a Trial Balance license
        $ILicenseApi.AssignLicensesToStaff($CreateApiContext, $NewStaff.StaffGuid.Guid.Guid, $TrialBalance.LicenseGuid.Guid.Guid)
    }
}
    
# ----------------------------------[  Working, and testing the following functions.]-------------------------------------
function Username-Check {
    <#
    This is a PowerShell function named "Username-Check" that generates a unique email by 
    concatenating the first character of the first name, the last name, and a domain name.
    It checks if the email is already taken in Active Directory and, if so, 
    tries again with a dot between the first name and last name. If both attempts fail, 
    it returns a warning message. The function takes two mandatory string parameters: "firstName" and "lastName".
    #>
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [String]$firstName,

        # Param2 help description
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false)]
        [String]$lastName
    )

    $domain = "@domain.com"
    $email = $firstName[0] + $lastName + $domain

    # Check for unique email in AD
    $Email1 = Get-ADObject -Filter { mail -eq $email }

    if ($Email1) {
        $email = $firstName + '.' + $lastName + $domain
        $Email2 = Get-ADObject -Filter { mail -eq $email }

        if ($Email2) {
            Write-Warning "Cannot generate a unique email for $firstName $lastName"
        }
        else {
            $email = $email
        }
    }
    else {
        $email = $email
    }

    $username, $domain = $email -split '@'

    $username
}

Function Get-Domain {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true)]
        [AllowEmptyString ()]
        [String]$NewHire
        
    )

    # Change the Domain base on new hire attribute
    switch ($NewHire) {
        "Something" { $Domain }
        default { $Domain = '@domain.com' }
    }

}

Function AssingEvolve-PhoneNumber {
    param($NewHire)

    # API endpoint URL
    $uri = "https://ossmosis.evolveip.net/api/as/provisioning/Submission"

    # API key or token for authentication
    # Define your credentials
    $username = "domain@domain.com"
    $password = "mLarg9lSv3x55S."
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $username, $password)))

    # Define the headers
    $headers = @{
        "Accept"        = "*/*"
        "Authorization" = "Basic $base64AuthInfo"
    }

    # Functions
    Function Select-FreeNumber {
        <#
   .SYNOPSIS
   This function generates a free number within a specified range for a given location.

   .PARAMETER NewHire
   The information for the new hire.

   .DESCRIPTION
   The function takes a 'NewHire' object as a parameter. 
   It defines an array of location data, each with location name, groupId, and a range of phone numbers.
   It then selects the data corresponding to the NewHire's location and randomly selects a phone number within that range.
   Finally, it creates and outputs a custom object containing the groupId and the selected phone number.
   #>
        param($NewHire)

        # Initialize the array
        $Table = @()

        # Depending on the location, add the corresponding custom object to the array
        switch ($NewHire.City) {
            "Bethesda" {
                $Table += [PSCustomObject]@{
                    Location         = 'lcation'
                    groupId          = '0001027323'
                    PhoneNumberRange = @{
                        Start = '+1-240'
                        End   = '+1-914'
                    }
                }
            }
            "Brooklyn" {
                $Table += [PSCustomObject]@{
                    Location         = 'lcation'
                    groupId          = '0001027459'
                    PhoneNumberRange = @{
                        Start = '+1-332'
                        End   = '+1-718'
                    }
                }
       
            }
            "Cranford" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001016606'
                    PhoneNumberRange = @{
                        Start = '+1-908'
                        End   = '+1-973'
                    }
                }
            }
            "Harrison" {
                $Table += [PSCustomObject]@{
                    Location         = 'lcation'
                    groupId          = '0001016600'
                    PhoneNumberRange = @{
                        Start = '+1-914'
                        End   = '+1-845'
                    }
                }
            }
            "Hauppauge" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001027324'
                    PhoneNumberRange = @{
                        Start = 'NA'
                        End   = '+1-631'
                    }
                }
            }
            "Middletown" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001031011'
                    PhoneNumberRange = @{
                        Start = 'NA'
                        End   = '+1-631'
                    }
                }
            }
            "NewburghBalmville" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001027325'
                    PhoneNumberRange = @{
                        Start = 'NA'
                        End   = '+1-845'
                    }
                }
            }
            "NewburghFostertown" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001016602'
                    PhoneNumberRange = @{
                        Start = 'NA'
                        End   = '+1-845'
                    }
                }
            }
            "ParkAvenue" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001027269'
                    PhoneNumberRange = @{
                        Start = 'NA'
                        End   = '+1-646'
                    }
                }
            }
            "Providence" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001027327'
                    PhoneNumberRange = @{
                        Start = 'NA'
                        End   = '+1-401'
                    }
                }
            }
            "Shelton" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001027328'
                    PhoneNumberRange = @{
                        Start = 'NA'
                        End   = '+1-203'
                    }
                }
            }
            "Stamford" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001016603'
                    PhoneNumberRange = @{
                        Start = '+1-203'
                        End   = '+1-475'
                    }
                }
            }
            "Wethersfield" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001016604'
                    PhoneNumberRange = @{
                        Start = '+1-203'
                        End   = '+1-860'
                    }
                }
            }
            "WoodcliffLake" {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001016608'
                    PhoneNumberRange = @{
                        Start = '+1-201'
                        End   = '+1-551'
                    }
                }
            }
            default {
                $Table += [PSCustomObject]@{
                    Location         = 'location'
                    groupId          = '0001016600'
                    PhoneNumberRange = @{
                        Start = '+1-914'
                        End   = '+1-845'
                    }
                }
            }
        }

        Function OSSmosisAssignedNumberList {

            # Define the URL
            $url = "https://ossmosis.evolveip.net/api/as/enterprise/eip-0001016600/OSSmosisAssignedNumberList"

            # Define the headers
            $headers = @{
                "Accept"        = "*/*"
                "Authorization" = "Basic $base64AuthInfo"
            }

            # Make the GET request
            $response = Invoke-RestMethod -Uri $url -Method GET -Headers $headers

            # Output the response
            $response.payload.groupAssignedEntityDTOList
        }
        $OSSmosisAssignedNumberList = OSSmosisAssignedNumberList
        $FreeNumber = ($OSSmosisAssignedNumberList | where { $_.status -like "open" } )
   
        # Retrieve the properties for the specified location
        $LocationProperty = $Table | where { ( $_.Location -eq $NewHire.City ) -or ($_.Location -eq "Harrison") }

        # Select a random phone number within the range specified for the location
        $Random = $FreeNumber | Where-Object { ($_.phoneNumber -like "$($LocationProperty.PhoneNumberRange.Start)*") -or ($_.phoneNumber -like "$($LocationProperty.PhoneNumberRange.End)*") } | Get-Random 

        # Create a custom object with the groupId and selected phone number
        $PSCustomObject = [PSCustomObject]@{
            groupId    = $LocationProperty.groupId
            FreeNumber = $Random.phoneNumber
        }
   
        # Return the custom object
        $PSCustomObject
    }
    $FreeNumber = Select-FreeNumber -NewHire $NewHire

    # Request body
    $payload = @{
        payload = @{
            activatePhoneNumbers     = $false 
            cleanupPreviousEntries   = $true
            description              = "ON_DEMAND user provisioning for gr-0001027323" 
            enterpriseId             = "eip-0001016600"
            groupId                  = ('gr-' + "$($FreeNumber.groupId)")
            platformIdentifier       = "broadsoft-f"  
            provisioningMode         = "ON_DEMAND"
            provisioningRequestType  = "net.evolveip.ossmosis.entities.provisioning.user.OssmosisUsersProvisioningRequestDTO"
            region                   = "US"
            requestHandlerClassNames = @(
                "net.evolveip.ossmosis.provisioning.request.handlers.users.OSSmosisUsersProvisioningAddRequestHandler"
            )
            scheduledRunTimeStamp    = $null
            stopOnPreviousFailure    = $false
            userId                   = "emial@domain.com"
            users                    = @(
                @{
                    deviceType             = "Microsoft Teams - Direct Routing"
                    invalid                = $false
                    sipRegistrar           = ""  
                    baseStationName        = ""
                    countryCode            = "US" 
                    customLineLabelId      = $null
                    emailAddress           = "$($NewHire.Email)"
                    enableCaribbeanDialing = $false
                    enableSMS              = $false
                    extension              = ""
                    firstName              = "$($NewHire.FirstName)"
                    internationalUser      = $false
                    lastName               = "$($NewHire.LastName)"
                    licenseType            = "EIPTEAMSVOICE"
                    macAddress             = ""
                    mobile                 = ""
                    phoneNumber            = $FreeNumber.FreeNumber 
                    stageAreaId            = 67950578
                    teamsDomain            = "c1016600.teams.evolveip.net"
                    timeZone               = @{
                        timeZone    = "America/New_York"
                        displayName = "(GMT-04:00) (US) Eastern Time"
                    }
                    userId                 = $NewHire.Username
                    vlanId                 = ""
                    yahooId                = ""
                }
            )
        }
    }

    # Send POST request
    $response = Invoke-RestMethod `
        -Uri $uri `
        -Method Post `
        -Body ($payload | ConvertTo-Json -Depth 10) `
        -ContentType "application/json" `
        -Headers $headers

    # Output response
    if ($response) {
        Grant-CsTenantDialPlan -Identity "$($NewHire.Email)" -PolicyName 'EvolveIP-TenantDialPlan'
        Grant-CsOnlineVoiceRoutingPolicy -Identity "$($NewHire.Email)" -PolicyName 'EvolveIP-East'
    }

    $NewHire | Add-Member -NotePropertyMembers @{'PhoneNumber' = $FreeNumber.FreeNumber }

}

Function new-ADUserWithHomeFolder {
    param($Username)

    $HomeFolderPath = "\\ha-file-01.odmd.local\redirectedfolders\$username"

    # Set the home directory path for the user
    $Null = Set-ADUser $UserName -HomeDirectory $HomeFolderPath -HomeDrive "Z:"

    # Create the user's home folder on the share

    $Null = New-Item -Path $HomeFolderPath -ItemType Directory -Force -ErrorAction SilentlyContinue
    #New-Item -Path "\\ha-file-01.odmd.local\redirectedfolders\$username\Documents" -ItemType Directory -Force -ErrorAction SilentlyContinue
    #New-Item -Path "\\ha-file-01.odmd.local\redirectedfolders\$username\Desktop" -ItemType Directory -Force -ErrorAction SilentlyContinue
    #New-Item -Path "\\ha-file-01.odmd.local\redirectedfolders\$username\Downloads" -ItemType Directory -Force -ErrorAction SilentlyContinue

    # Grant the user permission to their home folder
    $Acl = Get-Acl -Path $HomeFolderPath
    $UserSID = (Get-ADUser $UserName).SID
    $FileSystemAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($UserSID, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($FileSystemAccessRule)
    $Null = Set-Acl -Path $HomeFolderPath -AclObject $Acl


    "Creating Z folder to the user: $($NewHire.Username)"

}

Function Add-ADGroups {
    <#
    This function takes four mandatory parameters: $Username, $Office, $Title, and $Department. 
    It then queries Active Directory for users in a specific organizational unit, filters those users by office, 
    title, and department, groups the resulting AD groups by name and removes duplicates, 
    and finally adds the user to each of those AD groups. 
    The function also writes a message to the console indicating which AD groups were added to the user.
    #>

    # Define function parameters
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


    # Initialize ADusers as arrays
    $ADusers1 = @()
    $ADusers2 = @()
    $ADgroups = @()


    # >>>>>>>>>>>>>>>>>>>>>>>>>>> [ Title ] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    # Get all users in the specified OU with the required properties and filter them by title, and department
    $ADusers1 += Get-ADUser -Filter * `
        -SearchBase "OU=ODMD,DC=odmd,DC=local" `
        -Properties Department, title, office, MemberOf, Enabled  -Server 'HA-DC1' | 
    Where-Object { ($_.Department -eq $Department) -and 
                   ($_.title -eq $Title) -and 
                   ($_.office -eq $Office) -and 
                   ($_.Enabled) } 

    # Group the AD groups by name and filter out groups that the user is already a member of
    $ADgroups += foreach ($Group in ($ADusers1.MemberOf | group)) {
        if (($Group).Count -ge (($ADusers.Count - 1))) {
            $Group.Group | select -Unique
        }
    }

    # >>>>>>>>>>>>>>>>>>>>>>>>>>> [Description] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    # Get all users in the specified OU with the required properties and filter them by title, and department
    $ADusers2 += Get-ADUser -Filter * `
        -SearchBase "OU=ODMD,DC=odmd,DC=local" `
        -Properties Department, Description, office, MemberOf, Enabled  -Server 'HA-DC1' | 
    Where-Object { ($_.Department -like $Department) -and 
                   ($_.Description -like $Title) -and 
                   ($_.office -like $Office) -and 
                   ($_.Enabled) } 

    # Group the AD groups by name and filter out groups that the user is already a member of
    $ADgroups += foreach ($Group in ($ADusers2.MemberOf | group)) {
        if (($Group).Count -ge (($ADusers.Count - 1))) {
            $Group.Group | select -Unique
        }
    }     
    # >>>>>>>>>>>>>>>>>>>>>>>>>>> [] <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    $ADgroups = $ADgroups | select -Unique

    # Add the user to each AD group
    $ADgroups | Where-Object { $_ } | ForEach-Object { Add-ADGroupMember -Identity $_ -Members $Username -Server 'HA-DC1' } 

    if ($ADgroups) { 
        # for the moment, the ADGroups will be set to pending manual by default intill the function is more reliable 
        #  $NewHire | Add-Member -NotePropertyMembers @{'ADGroups' = 'Yes'; } 
    }

    # Write a message to the console indicating which AD groups were added to the user
    Write-Host "Adding AD groups to the user: $Username" -ForegroundColor Yellow
}

Function AD-changes {
    param($NewHire)

    $IndiaDenyGroup = (
        'Deny Access - HA-File-04 All E Drive but Audit India and PM',
        'Deny Access - TS-HA3',
        'Deny Access - TS-HA2',
        'Deny Access - Yearli',
        'Deny Access - HA-C1',
        'Deny Access - HA-VCenter6',
        'Deny Access - TS-HA1',
        'Deny Access - HA-S1',
        'Deny Access - PKFOD-01',
        'Ctx-Policies-OPSEU' ,
        'Deny Access - HA-SQL4 EFG Drives',
        'Deny Access - HA-FILE-05 E Drive',
        'Deny Access - HA-APPS-01 E Drive',
        'Deny Access - HA-APP1 E Drive',
        'Deny Access - HA-SQL2 E Drive',
        'Deny Access - HA-PROSTAFF C Drive',
        'Deny Access - FILE1',
        'Deny Access - GKG-APP1',
        'Deny Access - CRA-DC1 E Drive',
        'Deny Access - MAIN2015',
        'Deny Access - MBCT320',
        'Deny Access - WCL-DC1 D Drive',
        'Deny Access - WCL-DC DF Drives',
        'Deny Access - HA-SURE2 EF Drives',
        'Deny Access - HA-SURE EFGH Drives',
        'Deny Access - BETH-DC1 C Drive',
        'Deny Access - NTSERVER',
        'Deny Access - WE-DC1 E Drive',
        'Deny Access - PAR-DC1 E Drive',
        'Deny Access - C drive on HA-APPS',
        'Deny Access - HA-SQL3 E drive',
        'Deny Access - HA-File-03 E drive',
        'Deny Access - HA-Tax E drrive',
        'Deny Access - HA-File-02 E drive'
    )

    $DefaultGroups = (
        'AVD.US-Users',
        'Exclaimer PKFOD',
        'Group LastPass Users',
        'Group LastPass Mobile',
        'iChannel Users'
    )

    if ($NewHire.Status -eq "Contractor/Consultant") {
        foreach ($group in $IndiaDenyGroup) {
            Add-ADGroupMember -Identity $group `
                -Members $NewHire.Username `
                -Server 'HA-DC1' `
                -confirm:$false
        }
    }

    foreach ($group in $DefaultGroups) {
        Add-ADGroupMember -Identity $group `
            -Members $NewHire.Username `
            -Server 'HA-DC1' `
            -confirm:$false
    }

    if ($NewHire.Status -eq "Employee (Full or Part Time)") {
        Set-ADUser -Identity $NewHire.Username `
            -Server 'HA-DC1' `
            -Replace @{'extensionAttribute3' = $NewHire.CellPhoneNumber } `
            -OfficePhone $NewHire.CellPhoneNumber -ErrorAction SilentlyContinue
    }

    # Coppy-paste AD groups from users with the same Office, Jobtitle, Department
    Add-ADGroups -Username $NewHire.Username `
        -Office (AD-Office -Location $NewHire.City) `
        -Title $NewHire.JobTitle -Department $NewHire.Department -ErrorAction SilentlyContinue 

    
    # Add, and remove Mobile phone
    Manage-ADUserMobilePhone -UserPrincipalName $NewHire.Username -MobilePhone $NewHire.PersonalNumber -showNumber $NewHire.ShowNumber

    # Set SMTP Address
    $SMTPpkfod = ('SMTP:' + $NewHire.Username + '@PKFOD.COM')
    Set-ADUser -Identity $NewHire.Username -add @{"proxyaddresses" = $SMTPpkfod } -Server 'HA-DC1'

    # Static Groups
    #if($NewHire.City -eq ){}

    # Add the LogonScript
    if (!($NewHire.State -eq 'Massachusetts')) { Set-LogonScript -NewHire $NewHire }

    # Add Certifications to the AD account
    if ($NewHire.Certification) { Set-ADUser -Identity $NewHire.Username -add @{"ExtensionAttribute2" = $NewHire.Certification } -Server 'HA-DC1' }
   
}

# ----------------------------------[                                               ]-------------------------------------

<# New updates list for version 1.1
    Function updated 
    * Username-Check; maked so this function also lookis for contacts username. If a contact is using the username, firstname.lastname will be use.
    * Add-ADGroups; this function will also look for jobs titles in the description of users.
    * AD-Office; users from "South Africa", "Mumbai", "Philippines" will now say 'international' in the attribute in AD.
    * Set-LogonScript; users with 'international' in the office attribute in AD will now get assign the 'HRlogon-65.bat' logon script.

    New Function
    * AssingEvolve-PhoneNumber; this function uses API to connect to Evolve phone systems. it looks for free numbers base on the user's location.
      it then assigns those free numebrs.
    * new-ADUserWithHomeFolde; it create a folder in '\ha-file-01.odmd.local\redirectedfolders' with the permission for the user.
      it also updates the AD attibute Profile > home folder with the path.
#>


#test 

# ----------------------------------[     Controller     ]-------------------------------------#


Function PKFOD-NewHire {
    cls
    # Get a list of Engagement

    # Get the sharepoint list
    $PnPListItems = (Get-PnPListItem -List "New Hire Script Input").FieldValues 
    $Global:NewHireList = foreach ($Item in $PnPListItems) {
    
        $property = @{
            ID                = $Item.ID
            FirstName         = $Item.field_5
            LastName          = $Item.field_6
            Middle            = $Item.field_7
            StartDate         = ($Item.field_8.ToString("MM/dd/yyyy"))
            State             = $Item.field_9
            City              = $Item.City
            Status            = $item.field_18
            Department        = $Item.field_19
            JobTitle          = $Item.JobTitle
            ApprovalStatus    = $Item.ApprovalStatus
            Expedate          = $Item.Expedite
            PersonalNumber    = $Item.PhoneNumber
            ShowNumber        = $Item.ShowNumber
            EmployeeID        = $Item.EmployeeID
            Certification     = $Item.Certification
            ReturningEmployee = $Item.ReturningEmployee
        }
    
        $Object = New-Object -TypeName PSobject -Property $property
        $Object
    } 
    
    # Iterate through each user object in the $Global:NewHireList variable
    $Global:NewHireList = foreach ($user in $Global:NewHireList) {
        # Convert the StartDate property from a string to a datetime object
        $startDate = [datetime]$user.StartDate

        # Calculate the number of days between the current date and the StartDate
        $daysUntilStart = (New-TimeSpan -Start (Get-Date) -End $startDate).Days

        # Check if the StartDate is less than 5 days in the future and greater than or equal to the current date,
        # and the ApprovalStatus is "Pending to Process"
        if ((($daysUntilStart -ge 0) -and 
             ($daysUntilStart -lt 21) -and 
             ($user.ApprovalStatus -eq "Approved")) -or
             (($user.Expedate -eq "Yes") -and 
             ($user.ApprovalStatus -eq "Approved")) -and 
             (!($user.ReturningEmployee -eq "Yes"))) {
            # If the conditions are met, print the user object to the console
            $user
        }
    }
   

    # Get License count
    "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<[                 Geting License count                               ]>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    "; License-Count


    # Iterate through each user object in the $Global:NewHireList variable check for licenses 
    "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<[         Iterate through each user for available licenses           ]>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    "; $Global:NewHireList = foreach ($NewHire in $Global:NewHireList) { License-check -NewHireList $NewHire }


    # Create the AD Accounts
    "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<[               Creating the AD Accounts                             ]>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    "; $Global:NewHireList = foreach ($NewHire in $Global:NewHireList) {
        AD-Account -NewHire $NewHire; #Write-Host "The AD account has been created for $($NewHire.firstname)"
    }
    

    # Assign the users to Microsoft 365 E3 License, and Waiting for Mailboxes for be create
    "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<[ Waiting for AD account to be in the cloud  ]>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    "; foreach ($NewHire in $Global:NewHireList) { Waitfor-sycn -Email $NewHire.Email -status $NewHire.Status -Username $NewHire.Username }  
    

    "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<[              Waiting for Mailbox to be created                     ]>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    "; foreach ($NewHire in $Global:NewHireList) { Waitfor-Mailbox -Email $NewHire.Email -Username $Username }  

    # Create CreateApiContext
    $CreateApiContext = CreateApiContext

    foreach ($NewHire in $Global:NewHireList) {
        "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<[             AD Changes, Mailbox Changes, IC creation               ]>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    "
        # Change AD address, and front desk number base un office location - Disabled
        change-physicaladdress -Username $NewHire.Username

        # Set some mailbox settings - Disabled
        MailboxSettings -UserPrincipalName $NewHire.Email

        # Create IChannel account  
        if (!(($NewHire.Department -eq "Admin - Marketing") -or ($NewHire.Department -eq "Admin - Human Resources"))) {
            #   Ichannel-Account -cUserID $NewHire.Username -cLastName $NewHire.LastName -cFirstName $NewHire.FirstName -cEmail $NewHire.Email
        }

        # If the user is from Boston/woburm, Assing a phone number  - Disabled
        if (($NewHire.State -eq "Massachusetts") -and ($NewHire.Status -eq "Employee (Full or Part Time)")) {
            IF ($NewHire.Status -eq "Employee (Full or Part Time)") {
                #   Assign-PhoneNumber -FirstName $NewHire.FirstName -LastName $NewHire.LastName -Email $NewHire.Email -username $NewHire.Username
            }
        }

        # If the user is not from MA, and is a full time, assign evolve phone number.
        if ((!($NewHire.State -eq "Massachusetts")) -and ($NewHire.Status -eq "Employee (Full or Part Time)")) {
            "Create a phone number on Evolce"
            AssingEvolve-PhoneNumber -NewHire $NewHire
        }

        # Apply minor AD changes to the user
        AD-changes -NewHire $NewHire

        # Create Engagement account
        If (!($NewHire.Department -like "Admin*") -or ($NewHire.Department -like "H*")) {
            # "Creating Engagement account for $($NewHire.Username)"
            # New-EngagementUser -newHire $NewHire -CreateApiContext $CreateApiContext
        }

        # Create Home Drive folder
        if (!($NewHire.State -eq "Massachusetts")) { new-ADUserWithHomeFolder -Username $NewHire.Username }
         
        #  something something add later.
        $Null = Set-PnPListItem -List "New Hire Script Input" -Identity $newHire.ID -Values @{"ApprovalStatus" = "Pending Flow";
            "Username"                                                                                         = $NewHire.Email
        }
        $sharepointFlowList = Add-PnPListItem -List "New Hire Script Flow" -Values @{'NewHire' = $NewHire.Email 
            'Date'                                                                             = $NewHire.StartDate
            'HomeOffice'                                                                       = $NewHire.City
            'Department'                                                                       = $NewHire.Department
        }
        if ($NewHire.ADGroups -eq 'Yes') {
            $Null = Set-PnPListItem -List "New Hire Script flow" `
                -Identity  $sharepointFlowList.ID `
                -Values @{'ADGroups' = 'Completed' }
        }
    }

    #Send emails
    #  Email-Notification
}

Function Remove-MFAGroup {

    # Get the sharepoint list
    $PnPListItems = (Get-PnPListItem -List "New Hire Script Input").FieldValues 
    $MFANewHireList = foreach ($Item in $PnPListItems) {
    
        $property = @{
            ID        = $Item.ID
            StartDate = ($Item.field_8.ToString("MM/dd/yyyy"))
            Username  = $Item.Username
        }
    
        $Object = New-Object -TypeName PSobject -Property $property
        $Object
    } 

    foreach ($NewHire in $MFANewHireList) {
    
        # Check if Username property of NewHire object is not null or empty
        if (![string]::IsNullOrEmpty($NewHire.Username)) {
    
            # Convert StartDate to a DateTime object if it's not already
            $startDate = $NewHire.StartDate -as [DateTime]
    
            # Check if StartDate is not null and is equal to yesterday's date
            if (($startDate).AddDays(-2).ToString("MM/dd/yyyy") -eq (Get-Date).ToString("MM/dd/yyyy")) {
                
                # Remove the user to 'MFA Exclusion Group'
                Remove-AzureADGroupMember -ObjectId '28ade99f-e283-4442-96ec-2437279e8d5f' `
                    -MemberId (Get-AzureADUser -ObjectId $NewHire.Username).ObjectId

            }
        }   
    }

    $AzureG = "c29f8193-9b6f-4a40-935b-1e5c8697e8b5", '8ccc4f3f-beb7-45b2-9c79-4b5174ceef3b'
    foreach ($G in $AzureG) {

        # Retrieves all members of the specified Azure AD group
        $AzureADGroupMember = Get-AzureADGroupMember -ObjectId "$G" -All $true

        # For each user in the Azure AD group
        foreach ($Useremail in $AzureADGroupMember) {

            # Get the UserPrincipalName of the current user
            $email = $Useremail.UserPrincipalName
    
            # Get the AD user based on UserPrincipalName that are disabled and are in the 'Disabled' Organizational Unit
            $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$email'" | 
            where { ($_.Enabled -like "false") -and ($_.DistinguishedName -like "*Disabled*") }
        
            # If such a user exists
            if (($ADUser.Enabled -like "false") -and ($ADUser.DistinguishedName -like "*Disabled*") ) {

                # Remove the user from the 'Standard Azure group'
                Remove-AzureADGroupMember -ObjectId "$G" `
                    -MemberId (Get-AzureADUser -ObjectId $ADUser.UserPrincipalName).ObjectId
            }
            #>
        }
    }
}

#Run the scripts
PKFOD-NewHire
Remove-MFAGroup
Start-Sleep -Seconds (30)