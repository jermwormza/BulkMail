param(
    [string]$FromAddress = "",
    [string]$EmailListFilePath = "",
    [string]$EmailSubject = "",
    [string]$EmailBody = "",
    [string]$AttachmentFilePath = "",
    [bool]$UseAttachmentAsBody = $false,
    [bool]$Debug = $false,
    [string]$BccAddresses = "",
    [bool]$IncludeEmailListAsBcc = $false
)
# $Debug=$true
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$usedDialog = $false
$optionsFilePath = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) -ChildPath "bulkmail_options.xml"

# Function to write debug messages
function Write-DebugMessage {
    param (
        [string]$Message
    )
    if ($Debug) {
        Write-Host $Message
    }
}

# Function to update the command line preview
function Update-CommandLinePreview {
    $commandLine = "powershell.exe -File bulkmail.ps1 -FromAddress `"$($fromAddressTextbox.Text)`" -EmailListFilePath `"$($emailListFileTextbox.Text)`" -EmailSubject `"$($emailSubjectTextbox.Text)`" -EmailBody `"$($emailBodyTextbox.Text)`" -AttachmentFilePath `"$($attachmentFileTextbox.Text)`" -UseAttachmentAsBody:$($useAttachmentAsBodyCheckbox.Checked) -IncludeEmailListAsBcc:$($includeEmailListAsBccCheckbox.Checked)"
    $commandLineTextbox.Text = $commandLine
}

# Create form only if required parameters are not provided
if (-not ($FromAddress -and $EmailListFilePath -and $EmailSubject -and $EmailBody)) {
    # Load saved options if they exist and no command line options are provided
    if (-not ($FromAddress -or $EmailListFilePath -or $EmailSubject -or $EmailBody -or $AttachmentFilePath -or $BccAddresses -or $IncludeEmailListAsBcc) -and (Test-Path $optionsFilePath)) {
        Write-DebugMessage "Loading saved options from file: $optionsFilePath"
        $savedOptions = Import-Clixml -Path $optionsFilePath
        $FromAddress = $savedOptions.FromAddress
        $EmailListFilePath = $savedOptions.EmailListFilePath
        $EmailSubject = $savedOptions.EmailSubject
        $EmailBody = $savedOptions.EmailBody
        $AttachmentFilePath = $savedOptions.AttachmentFilePath
        $UseAttachmentAsBody = $savedOptions.UseAttachmentAsBody
        $BccAddresses = $savedOptions.BccAddresses
        $IncludeEmailListAsBcc = $savedOptions.IncludeEmailListAsBcc
    }

    $usedDialog = $true
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Bulk Email Sender"
    $form.Size = New-Object System.Drawing.Size(600, 420)  # Adjusted height to 450
    $form.StartPosition = "CenterScreen"

    # From Address
    $fromAddressLabel = New-Object System.Windows.Forms.Label
    $fromAddressLabel.Text = "From Address:"
    $fromAddressLabel.Location = New-Object System.Drawing.Point(10, 20)
    $fromAddressLabel.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($fromAddressLabel)

    $fromAddressTextbox = New-Object System.Windows.Forms.TextBox
    $fromAddressTextbox.Location = New-Object System.Drawing.Point(160, 20)
    $fromAddressTextbox.Size = New-Object System.Drawing.Size(300, 20)
    $fromAddressTextbox.Text = $FromAddress
    $fromAddressTextbox.Add_TextChanged({ Update-CommandLinePreview })
    $form.Controls.Add($fromAddressTextbox)

    # Email List File
    $emailListFileLabel = New-Object System.Windows.Forms.Label
    $emailListFileLabel.Text = "Email List File:"
    $emailListFileLabel.Location = New-Object System.Drawing.Point(10, 50)
    $emailListFileLabel.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($emailListFileLabel)

    $emailListFileTextbox = New-Object System.Windows.Forms.TextBox
    $emailListFileTextbox.Location = New-Object System.Drawing.Point(160, 50)
    $emailListFileTextbox.Size = New-Object System.Drawing.Size(300, 20)
    $emailListFileTextbox.Text = $EmailListFilePath
    $emailListFileTextbox.Add_TextChanged({ Update-CommandLinePreview })
    $form.Controls.Add($emailListFileTextbox)

    $emailListFileButton = New-Object System.Windows.Forms.Button
    $emailListFileButton.Text = "Browse..."
    $emailListFileButton.Location = New-Object System.Drawing.Point(470, 50)
    $emailListFileButton.Size = New-Object System.Drawing.Size(75, 23)
    $emailListFileButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $emailListFileTextbox.Text = $openFileDialog.FileName
            Update-CommandLinePreview
        }
    })
    $form.Controls.Add($emailListFileButton)

    $includeEmailListAsBccCheckbox = New-Object System.Windows.Forms.CheckBox
    $includeEmailListAsBccCheckbox.Text = "Use Email List as BCC"
    $includeEmailListAsBccCheckbox.Location = New-Object System.Drawing.Point(160, 80)
    $includeEmailListAsBccCheckbox.Size = New-Object System.Drawing.Size(200, 20)
    $includeEmailListAsBccCheckbox.Checked = $IncludeEmailListAsBcc
    $includeEmailListAsBccCheckbox.Add_CheckedChanged({ Update-CommandLinePreview })
    $form.Controls.Add($includeEmailListAsBccCheckbox)

    # Email Subject
    $emailSubjectLabel = New-Object System.Windows.Forms.Label
    $emailSubjectLabel.Text = "Email Subject:"
    $emailSubjectLabel.Location = New-Object System.Drawing.Point(10, 110)
    $emailSubjectLabel.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($emailSubjectLabel)

    $emailSubjectTextbox = New-Object System.Windows.Forms.TextBox
    $emailSubjectTextbox.Location = New-Object System.Drawing.Point(160, 110)
    $emailSubjectTextbox.Size = New-Object System.Drawing.Size(300, 20)
    $emailSubjectTextbox.Text = $EmailSubject
    $emailSubjectTextbox.Add_TextChanged({ Update-CommandLinePreview })
    $form.Controls.Add($emailSubjectTextbox)

    # Email Body
    $emailBodyLabel = New-Object System.Windows.Forms.Label
    $emailBodyLabel.Text = "Email Body:"
    $emailBodyLabel.Location = New-Object System.Drawing.Point(10, 140)
    $emailBodyLabel.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($emailBodyLabel)

    $emailBodyTextbox = New-Object System.Windows.Forms.TextBox
    $emailBodyTextbox.Location = New-Object System.Drawing.Point(160, 140)
    $emailBodyTextbox.Size = New-Object System.Drawing.Size(300, 100)
    $emailBodyTextbox.Multiline = $true
    $emailBodyTextbox.Text = $EmailBody
    $emailBodyTextbox.Add_TextChanged({ Update-CommandLinePreview })
    $form.Controls.Add($emailBodyTextbox)

    # Attachment File
    $attachmentFileLabel = New-Object System.Windows.Forms.Label
    $attachmentFileLabel.Text = "Attachment File (Optional):"
    $attachmentFileLabel.Location = New-Object System.Drawing.Point(10, 250)
    $attachmentFileLabel.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($attachmentFileLabel)

    $attachmentFileTextbox = New-Object System.Windows.Forms.TextBox
    $attachmentFileTextbox.Location = New-Object System.Drawing.Point(160, 250)
    $attachmentFileTextbox.Size = New-Object System.Drawing.Size(300, 20)
    $attachmentFileTextbox.Text = $AttachmentFilePath
    $attachmentFileTextbox.Add_TextChanged({ Update-CommandLinePreview })
    $form.Controls.Add($attachmentFileTextbox)

    $attachmentFileButton = New-Object System.Windows.Forms.Button
    $attachmentFileButton.Text = "Browse..."
    $attachmentFileButton.Location = New-Object System.Drawing.Point(470, 250)
    $attachmentFileButton.Size = New-Object System.Drawing.Size(75, 23)
    $attachmentFileButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "All files (*.*)|*.*"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $attachmentFileTextbox.Text = $openFileDialog.FileName
            Update-CommandLinePreview
        }
    })
    $form.Controls.Add($attachmentFileButton)

    $useAttachmentAsBodyCheckbox = New-Object System.Windows.Forms.CheckBox
    $useAttachmentAsBodyCheckbox.Text = "Use Attachment as Body"
    $useAttachmentAsBodyCheckbox.Location = New-Object System.Drawing.Point(160, 280)
    $useAttachmentAsBodyCheckbox.Size = New-Object System.Drawing.Size(200, 20)
    $useAttachmentAsBodyCheckbox.Checked = $UseAttachmentAsBody
    $useAttachmentAsBodyCheckbox.Add_CheckedChanged({ Update-CommandLinePreview })
    $form.Controls.Add($useAttachmentAsBodyCheckbox)

    # Command Line Preview
    $commandLineLabel = New-Object System.Windows.Forms.Label
    $commandLineLabel.Text = "Command Line Preview:"
    $commandLineLabel.Location = New-Object System.Drawing.Point(10, 310)
    $commandLineLabel.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($commandLineLabel)

    $commandLineTextbox = New-Object System.Windows.Forms.TextBox
    $commandLineTextbox.Location = New-Object System.Drawing.Point(160, 310)
    $commandLineTextbox.Size = New-Object System.Drawing.Size(385, 20)
    $commandLineTextbox.ReadOnly = $true
    $form.Controls.Add($commandLineTextbox)

    # Copy to Clipboard button
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy to Clipboard"
    $copyButton.Location = New-Object System.Drawing.Point(160, 340)
    $copyButton.Size = New-Object System.Drawing.Size(120, 23)
    $copyButton.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($commandLineTextbox.Text)
        [System.Windows.Forms.MessageBox]::Show("Command line copied to clipboard.", "Info")
    })
    $form.Controls.Add($copyButton)

    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(290, 340)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($okButton)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(370, 340)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($cancelButton)

    # Help button
    $helpButton = New-Object System.Windows.Forms.Button
    $helpButton.Text = "Help"
    $helpButton.Location = New-Object System.Drawing.Point(450, 340)
    $helpButton.Add_Click({
        $helpMessage = @"
Bulk Email Sender Help:

1. From Address:
   - The email address from which the emails will be sent.

2. Email List File:
   - The file containing the list of recipients. Supported formats are CSV and Excel (.xlsx).
   - The file must have the name and email address as the first two columns.

3. Use Email List as BCC:
   - If checked, the email list will be used as BCC recipients, and the 'From Address' will be used as the 'To' address.

4. Email Subject:
   - The subject of the email.

5. Email Body:
   - The body of the email. It can contain HTML code.
   - You can use {Name} in the body, which will be replaced with the recipient's name from the list file.

6. Attachment File (Optional):
   - An optional file to attach to the email.

7. Use Attachment as Body:
   - If checked, the content of the attachment file (if it is a text or HTML file) will be used as the email body.
   - Any {Name} in the attachment content will be replaced with the recipient's name from the list file.

8. Command Line Preview:
   - Shows the command line that can be used to run the script with the current settings without prompting.

9. Copy to Clipboard:
   - Copies the command line preview to the clipboard.

10. OK:
    - Saves the settings and starts sending emails.

11. Cancel:
    - Closes the application without sending emails.

12. Help:
    - Shows this help message.

The options are saved and reused on the next run if nothing is specified on the command line.

This application makes use of the Sendmail script to send emails. The Sendmail script must be in the same directory as this script.
Configuration and help for the Sendmail script can be found in the Sendmail script.

"@
        [System.Windows.Forms.MessageBox]::Show($helpMessage, "Help")
    })
    $form.Controls.Add($helpButton)

    # Update command line preview when the form is shown
    Update-CommandLinePreview

     # Show form
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Cancel) {
        exit
    }


    # Get input values
    $FromAddress = $fromAddressTextbox.Text
    $EmailListFilePath = $emailListFileTextbox.Text
    $EmailSubject = $emailSubjectTextbox.Text
    $EmailBody = $emailBodyTextbox.Text
    $AttachmentFilePath = $attachmentFileTextbox.Text
    $UseAttachmentAsBody = $useAttachmentAsBodyCheckbox.Checked
    $IncludeEmailListAsBcc = $includeEmailListAsBccCheckbox.Checked

    # Save options to file
    $options = [PSCustomObject]@{
        FromAddress = $FromAddress
        EmailListFilePath = $EmailListFilePath
        EmailSubject = $EmailSubject
        EmailBody = $EmailBody
        AttachmentFilePath = $AttachmentFilePath
        UseAttachmentAsBody = $UseAttachmentAsBody
        IncludeEmailListAsBcc = $IncludeEmailListAsBcc
    }
    $options | Export-Clixml -Path $optionsFilePath
    Write-DebugMessage "Options saved to file: $optionsFilePath"
}

# Debug statements
Write-DebugMessage "From Address: $FromAddress"
Write-DebugMessage "Email List File Path: $EmailListFilePath"
Write-DebugMessage "Email Subject: $EmailSubject"
Write-DebugMessage "Email Body: $EmailBody"
Write-DebugMessage "Attachment File Path: $AttachmentFilePath"
Write-DebugMessage "Use Attachment as Body: $UseAttachmentAsBody"
Write-DebugMessage "Include Email List as BCC: $IncludeEmailListAsBcc"

# Output command line for future use if dialog was used
if ($usedDialog) {
    $commandLine = "powershell.exe -File bulkmail.ps1 -FromAddress `"$FromAddress`" -EmailListFilePath `"$EmailListFilePath`" -EmailSubject `"$EmailSubject`" -EmailBody `"$EmailBody`" -AttachmentFilePath `"$AttachmentFilePath`" -UseAttachmentAsBody:$UseAttachmentAsBody -IncludeEmailListAsBcc:$IncludeEmailListAsBcc"
    Write-Host "To run this script without prompting in the future, use the following command line:"
    Write-Host $commandLine
}

# Read email list
Write-DebugMessage "Reading email list from file: $EmailListFilePath"
if ($EmailListFilePath -like "*.csv") {
    $emailList = Import-Csv -Path $EmailListFilePath
} elseif ($EmailListFilePath -like "*.xlsx") {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
    }
    $emailList = Import-Excel -Path $EmailListFilePath | Select-Object -Property *
}
Write-DebugMessage "Email list read successfully. Total recipients: $($emailList.Count)"

# Get the full path of the currently running script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$sendmailScript = Join-Path -Path $scriptPath -ChildPath "sendmail.ps1"

# Prepare BCC addresses if the option is selected
if ($IncludeEmailListAsBcc) {
    $bccAddresses = ""
}

# Send emails
foreach ($recipient in $emailList) {
    Write-DebugMessage "Sending email to: $($recipient.Email)"

    if ($IncludeEmailListAsBcc) {
        $bccAddresses = $recipient.Email
        # $toAddress = $FromAddress ## Not Needed
    } else {
        $toAddress = $recipient.Email
    }

    if ($UseAttachmentAsBody -and -not [string]::IsNullOrEmpty($AttachmentFilePath) -and (Test-Path $AttachmentFilePath)) {
        $attachmentContent = Get-Content -Path $AttachmentFilePath -Raw
        $body = $attachmentContent
        $AttachmentFilePath = $null  # Do not include attachment if used as body
    } else {
        $body = $EmailBody
    }

    Write-DebugMessage "Replacing placeholders in email body."
    Write-DebugMessage $recipient.PSObject.Properties
    foreach ($field in $recipient.PSObject.Properties) {
        Write-DebugMessage "Replacing: {$($field.Name)} with: $($field.Value)"
        $body = $body -replace "{$($field.Name)}", $field.Value
    }

    try {
        & $sendmailScript -from $FromAddress -to $toAddress -subject $EmailSubject -body $body -attachment $AttachmentFilePath -bcc $BccAddresses
        Write-DebugMessage "Email sent to: $($recipient.Email)"
    } catch {
        Write-DebugMessage "Failed to send email to: $($recipient.Email). Error: $_"
    }
    Start-Sleep -Seconds 2
}
Write-DebugMessage "All emails sent successfully."
