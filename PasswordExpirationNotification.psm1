# PasswordExpirationNotification.psm1

# Function to get credentials from file
function Get-PWENCredential {
    param (
        [string]$CredentialName,
        [string]$CredentialType
    )

    $credentialFolder = Join-Path -Path $PSScriptRoot -ChildPath "Credential"
    if (-not (Test-Path -Path $credentialFolder)) {
        New-Item -ItemType Directory -Path $credentialFolder | Out-Null
    }

    $credentialFile = Join-Path -Path $credentialFolder -ChildPath "$CredentialName.xml"

    if (Test-Path -Path $credentialFile) {
        $credential = Import-Clixml -Path $credentialFile
    } else {
        if ($CredentialType -eq "MicrosoftGraph") {
            $clientAppId = Read-Host "Enter Microsoft Graph Client App ID"
            $tenantId = Read-Host "Enter Microsoft Graph Tenant ID"
            $clientSecret = Read-Host "Enter Microsoft Graph Client Secret" -AsSecureString
            $credential = New-Object System.Management.Automation.PSCredential($clientAppId, $clientSecret)
            $credential | Add-Member -MemberType NoteProperty -Name TenantId -Value $tenantId
        } elseif ($CredentialType -eq "ActiveDirectory") {
            $username = Read-Host "Enter Active Directory Username"
            $password = Read-Host "Enter Active Directory Password" -AsSecureString
            $credential = New-Object System.Management.Automation.PSCredential($username, $password)
        } else {
            Write-Error "Invalid CredentialType specified"
            return $null
        }

        $credential | Export-Clixml -Path $credentialFile
    }

    return $credential
}

# Function to set credentials
function Set-PWENCredential {
    param (
        [string]$CredentialName,
        [string]$CredentialType
    )

    $credentialFolder = Join-Path -Path $PSScriptRoot -ChildPath "Credential"
    if (-not (Test-Path -Path $credentialFolder)) {
        New-Item -ItemType Directory -Path $credentialFolder | Out-Null
    }

    $credentialFile = Join-Path -Path $credentialFolder -ChildPath "$CredentialName.xml"

    if ($CredentialType -eq "MicrosoftGraph") {
        $clientAppId = Read-Host "Enter Microsoft Graph Client App ID"
        $tenantId = Read-Host "Enter Microsoft Graph Tenant ID"
        $clientSecret = Read-Host "Enter Microsoft Graph Client Secret" -AsSecureString
        $credential = New-Object System.Management.Automation.PSCredential($clientAppId, $clientSecret)
        $credential | Add-Member -MemberType NoteProperty -Name TenantId -Value $tenantId
    } elseif ($CredentialType -eq "ActiveDirectory") {
        $username = Read-Host "Enter Active Directory Username"
        $password = Read-Host "Enter Active Directory Password" -AsSecureString
        $credential = New-Object System.Management.Automation.PSCredential($username, $password)
    } else {
        Write-Error "Invalid CredentialType specified"
        return
    }

    $credential | Export-Clixml -Path $credentialFile
    Write-Host "Credential '$CredentialName' has been set successfully."
}

# Function to remove credentials
function Remove-PWENCredential {
    param (
        [string]$CredentialName
    )

    $credentialFolder = Join-Path -Path $PSScriptRoot -ChildPath "Credential"
    $credentialFile = Join-Path -Path $credentialFolder -ChildPath "$CredentialName.xml"

    if (Test-Path -Path $credentialFile) {
        Remove-Item -Path $credentialFile -Force
        Write-Host "Credential '$CredentialName' has been removed successfully."
    } else {
        Write-Warning "Credential '$CredentialName' not found."
    }
}

# Function to send password expiration notifications
function Send-PasswordExpirationNotifications {
    param (
        [string]$OuPath,
        [int]$DaysBeforeExpiration,
        [bool]$EnableEmail = $true,
        [bool]$OutputResults = $true,
        [bool]$EnableTranscription = $true,
        [string]$LogDirectory = "C:\Logs",
        [string]$EmailTemplate = ""
    )

    # Get Microsoft Graph credentials
    $mgGraphCredential = Get-PWENCredential -CredentialName "MicrosoftGraph" -CredentialType "MicrosoftGraph"

    # Get Active Directory credentials
    $adCredential = Get-PWENCredential -CredentialName "ActiveDirectory" -CredentialType "ActiveDirectory"

    # Start transcription if enabled
    if ($EnableTranscription) {
        $logFile = Join-Path -Path $LogDirectory -ChildPath "PasswordExpirationNotification_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        Start-Transcript -Path $logFile -Force
    }

    # Log the start of the script
    Write-Host "Starting Password Expiration Notification script at $(Get-Date)"

    # Connect to Microsoft Graph
    $secureClientSecret = $mgGraphCredential.Password
    $credential = New-Object System.Management.Automation.PSCredential($mgGraphCredential.UserName, $secureClientSecret)
    Connect-MgGraph -TenantId $mgGraphCredential.TenantId -ClientSecretCredential $credential

    # Log the successful connection to Microsoft Graph
    Write-Host "Successfully connected to Microsoft Graph"

    # Get users in the specified OU with expiring passwords
    $users = Get-MgUser -Filter "userPrincipalName -like '$OuPath/*'" -Property "userPrincipalName,mail,passwordPolicies" | Where-Object {
        $_.PasswordPolicies -notcontains "Disabled" -and
        $_.PasswordPolicies -notcontains "NeverExpires" -and
        (New-TimeSpan -Start (Get-Date) -End $_.PasswordExpirationDate).Days -le $DaysBeforeExpiration
    }

    # Log the number of users found with expiring passwords
    Write-Host "Found $($users.Count) users with expiring passwords"

    # Define the HTML email body
    $defaultHtmlBody = @"
<html>
<body>
    <h1>Password Expiration Notification</h1>
    <p>Dear [USER_NAME],</p>
    <p>Your password will expire in [DAYS_REMAINING] days. Please change your password before it expires.</p>
    <p>Thank you,</p>
    <p>Your IT Department</p>
</body>
</html>
"@

    # Use EmailTemplate if specified and file exists, otherwise use defaultHtmlBody
    if ($EmailTemplate -and (Test-Path -Path $EmailTemplate)) {
        $htmlBody = Get-Content -Path $EmailTemplate -Raw
    } else {
        $htmlBody = $defaultHtmlBody
        if ($EmailTemplate) {
            Write-Warning "Email template file '$EmailTemplate' not found. Using default template."
        }
    }

    # Initialize an array to store the results
    $results = @()

    # Send email notifications to users with expiring passwords and collect results
    if ($EnableEmail) {
        foreach ($user in $users) {
            $daysRemaining = (New-TimeSpan -Start (Get-Date) -End $user.PasswordExpirationDate).Days
            $userHtmlBody = $htmlBody.Replace("[USER_NAME]", $user.DisplayName).Replace("[DAYS_REMAINING]", $daysRemaining)

            $emailParams = @{
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $user.Mail
                        }
                    }
                )
                Subject = "Password Expiration Notification"
                Body = @{
                    ContentType = "HTML"
                    Content = $userHtmlBody
                }
            }

            Send-MgUserMail -UserId $user.Id -Message $emailParams

            # Log the successful sending of the email
            Write-Host "Sent password expiration notification to $($user.Mail)"

            # Add user information to the results array
            $results += [PSCustomObject]@{
                SamAccountName = $user.UserPrincipalName.Split('@')[0]
                Email = $user.Mail
                DaysUntilExpiration = $daysRemaining
            }
        }
    } else {
        Write-Host "Email functionality is disabled. No emails will be sent."
    }

    # Output results if enabled
    if ($OutputResults) {
        Write-Host "Outputting results:"
        $results
    }

    # Log the end of the script
    Write-Host "Finished Password Expiration Notification script at $(Get-Date)"

    # Stop transcription if enabled
    if ($EnableTranscription) {
        Stop-Transcript
    }

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph
}


# Export module members
Export-ModuleMember -Function Get-PWENCredential, Set-PWENCredential, Remove-PWENCredential, Send-PasswordExpirationNotifications
