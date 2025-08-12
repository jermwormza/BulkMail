# This script sends an email using the Microsoft Graph API and the Send-EMailMessage from the module Mailozaurr.
# Mailozaurr is a PowerShell module that provides a simple way to send emails using the Microsoft Graph API.
# The module can be installed from the PowerShell Gallery:
# Install the required module by running the following command:
# Install-Module -Name Mailozaurr

param(
    [String]$Body = "This is a test email.",
    [String]$Subject = "Test Email " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
    [String[]]$To = "nobody@nowhere.com",
    [String]$From = "nobody@nowhere.com",
    [String[]]$Attachments = "",
    [bool]$Debug = $false,
    [String]$Bcc = ""
)

# Function to write debug messages
function Write-DebugMessage {
    param (
        [string]$Message
    )
    if ($Debug) {
        Write-Host $Message
    }
}

$computerName = $env:COMPUTERNAME
$userName = $env:USERNAME
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$credentialFilePath = Join-Path -Path $scriptPath -ChildPath "$computerName-$userName-Credential.xml"

Write-DebugMessage "Checking for credential file at: $credentialFilePath"
$promptForCredentials = $true
if (Test-Path $credentialFilePath) {
    Write-DebugMessage "Credential file found. Importing credentials."
    $credentials = Import-CliXml -Path $credentialFilePath
    if ($credentials.AppId -and $credentials.TenantId -and $credentials.Secret) {
        $appId = $credentials.AppId
        $tenantId = $credentials.TenantId
        $secureSecret = $credentials.Secret
        $promptForCredentials = $false
    }
}

if ($promptForCredentials) {
    Write-DebugMessage "Credential file not found or incomplete. Prompting for App ID, Tenant ID, and Secret value."
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enter Credentials"
    $form.Width = 400
    $form.Height = 300
    $form.StartPosition = "CenterScreen"

    $labels = @("App ID:", "Tenant ID:", "Secret Value:")
    $yPos = 20
    $textboxes = @()
    foreach ($labelText in $labels) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelText
        $label.AutoSize = $true
        $label.Location = New-Object System.Drawing.Point(10, $yPos)
        $form.Controls.Add($label)

        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = New-Object System.Drawing.Point(100, $yPos)
        $textbox.Width = 250
        if ($labelText -eq "Secret Value:") {
            $textbox.UseSystemPasswordChar = $true
        }
        $form.Controls.Add($textbox)
        $textboxes += $textbox
        $yPos += 30
    }

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(150, 200)
    $okButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($okButton)

    $helpButton = New-Object System.Windows.Forms.Button
    $helpButton.Text = "Help"
    $helpButton.Location = New-Object System.Drawing.Point(250, 200)
    $helpButton.Add_Click({
        $helpMessage = @"
Instructions to create and get Application AppID and Tenant ID for Mailozaurr:
1. Go to the Azure portal (https://portal.azure.com/).
2. In the search bar at the top, type 'Azure Active Directory' and select it from the results.
   - 'Manage Microsoft Entra ID' has replaced 'Azure Active Directory'.
3. Under 'Manage', select 'App registrations'.
4. Click on 'New registration' at the top.
5. Fill in the required fields and click 'Register'.
6. After registration, you will be taken to the app's overview page.
7. Copy the 'Application (client) ID' - this is your AppID.
8. Copy the 'Directory (tenant) ID' - this is your Tenant ID.
9. Under 'Manage', select 'Certificates & secrets'.
10. Click on 'New client secret' to generate a new secret value.
11. Copy the secret value and store it securely. This will be used as the Client Secret.
    - Do not store the Client Secret in plain text in your scripts. The script will encrypt the secret and store it in a file.
    - You may store the Client Secret in a secure vault or key management system.
    - If you lose the secret, you can generate a new one following steps 9 and 10.

Sendmail Script Help:

1. From Address:
   - The email address from which the emails will be sent.

2. To Addresses:
   - The email addresses of the recipients.

3. Email Subject:
   - The subject of the email.

4. Email Body:
   - The body of the email. It can contain HTML code.

5. Attachments (Optional):
   - An optional file to attach to the email.

6. BCC Addresses:
   - BCC email addresses.

7. Debug:
   - If set to true, debug messages will be displayed.

Note: The credentials are saved and reused on the next run if the credential file is found.
"@
        [System.Windows.Forms.MessageBox]::Show($helpMessage, "Help")
    })
    $form.Controls.Add($helpButton)

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $appId = $textboxes[0].Text
        $tenantId = $textboxes[1].Text
        $secretValue = $textboxes[2].Text
    } else {
        throw "Credential input was cancelled."
    }

    $secureSecret = ConvertTo-SecureString -String $secretValue -AsPlainText -Force | ConvertFrom-SecureString
    $credentials = [PSCustomObject]@{
        AppId = $appId
        TenantId = $tenantId
        Secret = $secureSecret
    }
    $credentials | Export-CliXml -Path $credentialFilePath
    Write-DebugMessage "Credential file created at: $credentialFilePath"
}

Write-DebugMessage "Importing secret from credential file."
$Credential = ConvertTo-GraphCredential -ClientID $appId -ClientSecretEncrypted $secureSecret -DirectoryID $tenantId
Write-DebugMessage "Credential imported successfully."

#log set up
$logfile = Join-Path -Path $scriptPath -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + ".log")

$EmailParams = @{
    From = $From
    To = $To
    Subject = $Subject
    Body = $Body
    Graph = $true
    Credential = $Credential
    Attachments = $Attachments
    Bcc = $Bcc
}

Send-EMailMessage @EmailParams
