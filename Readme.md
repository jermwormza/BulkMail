# Bulk Email Sender

## Overview

The Bulk Email Sender is a PowerShell script that allows you to send bulk emails using a list of recipients from a CSV or Excel file. It uses the Microsoft Graph API and the `Mailozaurr` module to send emails. It also has some limited 'mail merge' functionality where the specified body or a html/text attachment is used as the body coded with fields as {fieldname} and there is a matching fieldname in the recipients csv or excel file it will replace the placeholders with the data from the file. The Email column may contain multiple email addresses separated by a comma but there is no validation if addresses are valid.

The related sendmail script uses Office 365/Azure Graph functionality to send emails instead of smtp which adds an extra level of security and identity protection. If these are not available to you, you may replace the sendmail.ps1 with a different script as long as the parameter names remain the same. Outgoing eMails will be in the sent Items of the sending account. The from address should be a mailbox and the user sending should have "Send As" permissions (I haven't fully tested this yet, there may be some limitations).

There is no guarantee this works in all circumstances, it works for me in my environment. If you would like to contribute improvements, *ask*.

## Prerequisites

1. **Mailozaurr Module**: Install the required module by running the following command:

   ```powershell
   Install-Module -Name Mailozaurr
   Install-Module -Name Microsoft.Graph
   ```

2. **Azure Application Registration**:
   - Go to the Azure portal (<https://portal.azure.com/>).
   - In the search bar at the top, type 'Azure Active Directory' and select it from the results.
     - 'Manage Microsoft Entra ID' has replaced 'Azure Active Directory'.
   - Under 'Manage', select 'App registrations'.
   - Click on 'New registration' at the top.
   - Fill in the required fields and click 'Register'.
   - After registration, you will be taken to the app's overview page.
   - Copy the 'Application (client) ID' - this is your AppID.
   - Copy the 'Directory (tenant) ID' - this is your Tenant ID.

3. **Add API Permissions (Mail.Send and Mail.ReadWrite)**:
   - Under 'Manage', select 'API permissions'.
   - Click on 'Add a permission'.
   - Select 'Microsoft Graph'.
   - Choose Application permissions'.
   - In the search box, type 'Mail'.
   - Check the boxes for 'Mail.Send' and 'Mail.ReadWrite'.
   - Click 'Add permissions'.

4. **Grant Admin Consent**:
   - After adding the permissions, you will see them listed under 'Configured permissions'.
   - Click the 'Grant admin consent for [Your Organization]' button (It shows as a link above the permissions list).
   - Confirm the action when prompted. The status should update to 'Granted for [Your Organization]'.

## Bulk Email Sender Script (`bulkmail.ps1`)

### Bulkmail Parameters

- `FromAddress`: The email address from which the emails will be sent.
- `EmailListFilePath`: The file containing the list of recipients. Supported formats are CSV and Excel (.xlsx). The file must include a header row with at least `Name` and `Email` fields.
- `EmailSubject`: The subject of the email.
- `EmailBody`: The body of the email. It can contain HTML code.
- `AttachmentFilePath`: An optional file to attach to the email. If the attachment is a text or HTML file, placeholders in the format `{fieldname}` will be replaced with the corresponding data from the recipient list.
- `UseAttachmentAsBody`: If set to True, the content of the attachment file (if it is a text or HTML file) will be used as the email body.
- `Debug`: If set to true, debug messages will be displayed.
- `BccAddresses`: BCC email addresses.
- `IncludeEmailListAsBcc`: If set to true, the email list will be used as BCC recipients, and the 'From Address' will be used as the 'To' address.

### Bulkmail Usage

1. Run the script:

   ```powershell
   .\bulkmail.ps1
   ```

2. If required parameters are not provided, a form will be displayed to input the necessary information.

3. The options are saved and reused on the next run if nothing is specified on the command line. The saved options file is encrypted and is specific to each Computer and User it is run under.

4. Any fields in the Email List file will be available to be replaced in the body text or in the attachment html/txt file using {field name}.

### Bulkmail Help

The help button in the form provides detailed information about each field and the functionality of the script.

## Sendmail Script (`sendmail.ps1`)

### Sendmail Parameters

- `Body`: The body of the email. It can contain HTML code.
- `Subject`: The subject of the email.
- `To`: The email addresses of the recipients.
- `From`: The email address from which the emails will be sent.
- `Attachments`: An optional file to attach to the email. If the attachment is a text or HTML file, placeholders in the format `{fieldname}` will be replaced with the corresponding data from the recipient list.
- `Debug`: If set to true, debug messages will be displayed.
- `Bcc`: BCC email addresses.

### Sendmail Usage

1. Run the script:

   ```powershell
   .\sendmail.ps1
   ```

2. If the credential file is not found or incomplete, a form will be displayed to input the necessary credentials (App ID, Tenant ID, and Secret value).

3. The credentials are saved and reused on the next run if the credential file is found.
   1. The credentials file is encrypted and can only be used by the same user on the same computer.
   2. The credentials file name contains the Computer Name and User Name so the same folder can be ued by multiple users on multiple computers.
   3. You cannot create a credential file for a different user. If you need to use a scheduled task service account, you must first login interactively using that account and run the script to create the credentials file, after the file is there the script can be used non-interactively.
   4. All this means that you may need a secret for each user you want to enable to do emailing using this script. You can add many secrets to the registered app, I haven't needed many so not sure if there are any limits.

### Sendmail Help

The help button in the form provides detailed instructions to create and get the Application AppID and Tenant ID for Mailozaurr.

## Note

Ensure that the `sendmail.ps1` script is in the same directory as the `bulkmail.ps1` script for the Bulk Email Sender to function correctly. All configuration and credential files will be saved in the same folder as the script.
