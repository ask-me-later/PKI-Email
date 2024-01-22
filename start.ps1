############################################
#                                          #
#     Certificate issuer & e-mail sender   #
#        Developed by Bakuła, Kamil        #
#                                          #
############################################

#Please compile this script to an .EXE file using PS2EXE-GUI and use it as an desktop application.

Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
[System.Windows.Forms.Application]::EnableVisualStyles()

# Creating main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Certificate Issuer <custom name>"
$form.Size = New-Object System.Drawing.Size(550,550)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $False
$form.FormBorderStyle = "FixedSingle"

# Varbiables used for validation
$Global:CertificateFileTRUE = $False
$Global:CertificateSelectedTRUE = $False
$Global:TicketNumber = ""
$Global:selectedFilePath = ""
$Global:emailaddress = ""

# This base64 string holds the bytes that make up the icon (32x32 pixel image)
$iconBase64      = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAACTSURBVDhPvc7BCYNAEIXhLcS9iJBm0pN12ISQg0WkhDSQQyA1rDzJjOOb2c0hxMMHovN+TNPjVeDdDU1yx7YAH5fxoux7HoMLbMP5vvsSOQTcOIj8FOAxnPsH1QBUI40xuIBGPlpj0ICwIYvvRBjob1cd4plvLBcAjKzoRoQBqAWWZz7QAH+wAeDvIkUvefzHQC4rFpTdcLBpEc8AAAAASUVORK5CYII='
$iconBytes       = [Convert]::FromBase64String($iconBase64)

# initialize a Memory stream holding the bytes
$stream          = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
$Form.Icon       = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

# Create the menu strip
$stripMenu = New-Object System.Windows.Forms.MenuStrip
$stripMenu.Dock = "top"

# Create the "File" menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

# Create a context menu for the "Refresh" option
$refreshMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Create the "Refresh" menu item and define its click event
$refreshMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$refreshMenuItem.Text = "Refresh"
$refreshMenuItem.Add_Click({

    Try {
    # Handle the "Refresh" menu item click event
    $textbox0Task.Text = "RITM0X0X0X0"
    $Global:TicketNumber = ""
    $Global:selectedFilePath = ""
    $Global:emailaddress = ""
    $textbox0CSR.Text = "-----BEGIN CERTIFICATE REQUEST-----
[YOUR CSR text here]
-----END CERTIFICATE REQUEST-----"
    $button0GenCert.Enabled = $False
    $button1SelectCert.Enabled = $true
    $Global:selectedFilePath = $null
    $CustomFontLabel1CertPathEMPTY = New-Object System.Drawing.Font($label1CertPath.Font.FontFamily, $label1CertPath.Font.Size, [System.Drawing.FontStyle]::Regular)
    $label1CertPath.ForeColor = [System.Drawing.Color]::Black
    $label1CertPath.Font = $CustomFontLabel1CertPathEMPTY
    $label1CertPath.Text = "EMPTY"
    $textbox1Email.Text = "test@domain.com"
    $Global:CertificateSelectedTRUE = $False
    $Global:CertificateFileTRUE = $False
    $button1SendCert.Enabled = $False
    [System.Windows.Forms.MessageBox]::Show("All variables were reset.", "Refresh successful" , 0, "Warning")
    }

    Catch {
        [System.Windows.Forms.MessageBox]::Show("Unspecified error.`n`r`nPlease report to your administrator.", "Refresh error" , 0, "Error")
    }
})
# Add the "Refresh" menu item to the context menu
[void]$refreshMenu.Items.Add($refreshMenuItem)
# Assign the context menu to the "File" menu
$fileMenu.DropDown = $refreshMenu
[void]$stripMenu.Items.Add($fileMenu)

# Create the "About" menu
$aboutMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenu.Text = "Help"
[void]$stripMenu.Items.Add($aboutMenu)

# Add click event handlers for the menu items
$aboutMenu.Add_Click({ [System.Windows.Forms.MessageBox]::Show("Certificate issuer for <custom name>.`n`r`nPlease always run as admin. `n`r`nCertificate will be issued using: `n`r<your template name> template. `n`r<FQDN of SubCA> SubCA machine. `n`r<CA instance name>.`n`r`nThe certificate can also be sent by e-mail as a .zip attachment.`n`rE-mail message includes: `n`r-Attachment: Certificate file packed to a [Ticket-Number].zip format archive. `n`r-CC: Wintel group e-mail address. `n`r-Subject: Provided request number. `n`r`nThe ZIP file and CSR will always be stored under C:\Temp\ folder. In case of issues check if such folder exists. `n`r`nVersion 1.0 | January 2024`n`nDeveloped by Bakula, Kamil ~ Poland Lodz.", "Info" , 0, "Information") })

# Add the menu strip to the form
[void]$form.Controls.Add($stripMenu)

# Create Header name for Task section
$label0Task = New-Object System.Windows.Forms.Label
$label0Task.Location = New-Object System.Drawing.Point(10,40)
$label0Task.Size = New-Object System.Drawing.Size(140,20)
$label0Task.Text = "Service Request Number:"
$label0Task.Enabled = $true
[void]$form.Controls.Add($label0Task)

# Create Textbox for ticket number
$textbox0Task = New-Object System.Windows.Forms.TextBox
$textbox0Task.Location = New-Object System.Drawing.Point(10,65)
$textbox0Task.Size = New-Object System.Drawing.Size(300,20)
$textbox0Task.Enabled = $true
$textbox0Task.Text = "RITM0X0X0X0"
$textbox0Task.MaxLength = 20
[void]$form.Controls.Add($textbox0Task)

# Create Header name for CSR section
$label0Name = New-Object System.Windows.Forms.Label
$label0Name.Location = New-Object System.Drawing.Point(10,95)
$label0Name.Size = New-Object System.Drawing.Size(250,20)
$label0Name.Text = "Certificate Signing Request (CSR) text input:"
$label0Name.Enabled = $true
[void]$form.Controls.Add($label0Name)

# Create Textbox for CSR
$textbox0CSR = New-Object System.Windows.Forms.TextBox
$textbox0CSR.Location = New-Object System.Drawing.Point(30,120)
$textbox0CSR.Size = New-Object System.Drawing.Size(475,230)
$textbox0CSR.Multiline = $true
$textbox0CSR.Enabled = $true
$textbox0CSR.Name = "CSR"
$textbox0CSR.Text = "-----BEGIN CERTIFICATE REQUEST-----
[YOUR CSR text here]
-----END CERTIFICATE REQUEST-----"

# Create a ToolTip control for CSR text
$toolTipNameNotEmpty = New-Object Windows.Forms.ToolTip
$toolTipNameNotEmpty.IsBalloon = $true

$textbox0CSR.Add_TextChanged({
    # Check if the name is not empty
    if ($textbox0CSR.Text -like "") {
        $textbox0CSR.BackColor = [System.Drawing.Color]::LightPink
        # Display a tooltip near the textbox
        $toolTipNameNotEmpty.SetToolTip($textbox0CSR, "CSR file cannot be empty!")
        $button0GenCert.Enabled = $False

    }
    else {
        $textbox0CSR.BackColor = [System.Drawing.Color]::White
        $toolTipNameNotEmpty.Active = $False
        $button0GenCert.Enabled = $true
        }
})
[void]$form.Controls.Add($textbox0CSR)

# Create Button to generate certificate
$button0GenCert = New-Object System.Windows.Forms.Button
$button0GenCert.Location = New-Object System.Drawing.Point(185,360)
$button0GenCert.Size = New-Object System.Drawing.Size(150,20)
$button0GenCert.Enabled = $False
$button0GenCert.Text = "Generate Certificate"
$button0GenCert.Add_Click({

    Try {

    # Cert auth names here
    $caConfig = "<hostname of a CA>\<Instance name of CA>"

    # Cert template name here
    $templateName = "<your certificate template name>"

    # Assigning text of CSR textbox to a variable
    $csrText = $textbox0CSR.Text

    # Input of CSR textbox to a new .csr file under C:\Temp
    $csrFile = "C:\Temp\csrFile_Script.csr"
    Set-Content -Path $csrFile -Value $csrText

    # Submit the CSR and generate certificate
    $certRequest = certreq -submit -attrib "CertificateTemplate:$templateName" -config $caConfig $csrFile

    [System.Windows.Forms.MessageBox]::Show("Workflow completed. `n`rCertificate was generated.", "Info" , 0, "Info")
    }

    Catch {

    
    [System.Windows.Forms.MessageBox]::Show("Certificate was not generated. `n`rPlease contact your administrator.", "Error" , 0, "Error")
    }
    


})
[void]$form.Controls.Add($button0GenCert)

# Create separator for Validation section
$label2line = New-Object System.Windows.Forms.Label
$label2line.Location = New-Object System.Drawing.Point(10,385)
$label2line.Size = New-Object System.Drawing.Size(550,20)
$label2line.Text = "___________________________________________________________________________________"
$label2line.Enabled = $False
[void]$form.Controls.Add($label2line)

# Create header name for Validation section
$label2Valid = New-Object System.Windows.Forms.Label
$label2Valid.Location = New-Object System.Drawing.Point(150,410)
$label2Valid.Size = New-Object System.Drawing.Size(300,20)
$label2Valid.Text = "E-mail a certificate to requester (optional):"
$label2Valid.Enabled = $true
[void]$form.Controls.Add($label2Valid)

# Create Textbox for e-mail address
$textbox1Email = New-Object System.Windows.Forms.TextBox
$textbox1Email.Location = New-Object System.Drawing.Point(145,435)
$textbox1Email.Size = New-Object System.Drawing.Size(225,20)
$textbox1Email.Enabled = $true
$textbox1Email.Text = "test@domain.com"

# Create tooltip for e-mail address empty or invalid
$toolTipEmailNotEmpty = New-Object Windows.Forms.ToolTip
$toolTipEmailNotEmpty.IsBalloon = $true
$toolTipEmailNotAT = New-Object Windows.Forms.ToolTip
$toolTipEmailNotAT.IsBalloon = $true


# Check if e-mail was entered correctly
$textbox1Email.Add_TextChanged({
    if ($textbox1Email.Text -eq "") {
        # Text box is empty
        $textbox1Email.BackColor = [System.Drawing.Color]::LightPink
        $toolTipEmailNotEmpty.SetToolTip($textbox1Email, "E-mail cannot be empty!")
        $button1SendCert.Enabled = $False
        $Global:CertificateSelectedTRUE = $False
    }
    elseif ($textbox1Email.Text -notlike "*@*") {
        # Text box does not contain "@"
        $textbox1Email.BackColor = [System.Drawing.Color]::LightPink
        $toolTipEmailNotEmpty.SetToolTip($textbox1Email, "Invalid e-mail format!")
        $button1SendCert.Enabled = $False
        $Global:CertificateSelectedTRUE = $False
    }
    else {
        # Text box contains text and "@"
        $textbox1Email.BackColor = [System.Drawing.Color]::White
        $toolTipEmailNotEmpty.SetToolTip($textbox1Email, "") # Clear tooltip
        $Global:CertificateSelectedTRUE = $true

        if (($Global:CertificateFileTRUE -eq $true) -and ($Global:CertificateSelectedTRUE -eq $true)) {

        $button1SendCert.Enabled = $true

        }

        else {

        $button1SendCert.Enabled = $False

        }

    }
})

[void]$form.Controls.Add($textbox1Email)

# Create header for certificate file path
$label1SelectedCert = New-Object System.Windows.Forms.Label
$label1SelectedCert.Location = New-Object System.Drawing.Point(10,470)
$label1SelectedCert.Size = New-Object System.Drawing.Size(105,20)
$label1SelectedCert.Text = "Selected certificate:"
$label1SelectedCert.Enabled = $true
[void]$form.Controls.Add($label1SelectedCert)

# Create variable text for certificate file path
$label1CertPath = New-Object System.Windows.Forms.Label
$label1CertPath.Location = New-Object System.Drawing.Point(120,470)
$label1CertPath.Size = New-Object System.Drawing.Size(500,20)
$label1CertPath.Text = "EMPTY"
$label1CertPath.Enabled = $true
[void]$form.Controls.Add($label1CertPath)

# Create Button to select certificate to send
$button1SelectCert = New-Object System.Windows.Forms.Button
$button1SelectCert.Location = New-Object System.Drawing.Point(75,495)
$button1SelectCert.Size = New-Object System.Drawing.Size(150,20)
$button1SelectCert.Enabled = $true
$button1SelectCert.Text = "Select certificate"
$button1SelectCert.Add_Click({

    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = 'All files (*.*)|*.*'  # Filter for file types
    $result = $fileDialog.ShowDialog()

    # Check if the user selected a file
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        # Store the file path in a variable
        $Global:selectedFilePath = $fileDialog.FileName
        $label1CertPath.Text = $Global:selectedFilePath
        $CustomFontLabel1CertPathOK = New-Object System.Drawing.Font($label1CertPath.Font.FontFamily, $label1CertPath.Font.Size, [System.Drawing.FontStyle]::Bold)
        $label1CertPath.Font = $CustomFontLabel1CertPathOK
        $label1CertPath.ForeColor = [System.Drawing.Color]::Green
        $label1CertPath.Font.Bold
        $Global:CertificateFileTRUE = $true
        if (($Global:CertificateFileTRUE -eq $true) -and ($Global:CertificateSelectedTRUE -eq $true)) {

        $button1SendCert.Enabled = $true

        }

        else {

        $button1SendCert.Enabled = $False

        }
    }

})
[void]$form.Controls.Add($button1SelectCert)

# Create Button to send a cert
$button1SendCert = New-Object System.Windows.Forms.Button
$button1SendCert.Location = New-Object System.Drawing.Point(290,495)
$button1SendCert.Size = New-Object System.Drawing.Size(150,20)
$button1SendCert.Enabled = $False
$button1SendCert.Text = "Send certificate"
$button1SendCert.Add_Click({
    
    Try {

    $Global:TicketNumer = $textbox0Task.Text

    Compress-Archive -Path $Global:selectedFilePath -DestinationPath C:\Temp\$Global:TicketNumer -Force

    $Checkifzipcreated = $(Try { Get-Item C:\Temp\$Global:TicketNumer.zip } Catch {$null})
    If ($Checkifzipcreated -eq $null) {

    [System.Windows.Forms.MessageBox]::Show("$Global:TicketNumer.zip could not be created! `n`rCheck if you are running script as admin and if specified files exist.", "Error" , 0, "Error")
    [System.Windows.Forms.MessageBox]::Show("The E-mail was not sent.", "Error" , 0, "Error")
    }

    else {
    [System.Windows.Forms.MessageBox]::Show("ZIP archive $Global:TicketNumer.zip has been created containing the file $Global:selectedFilePath. `n`rPreparing e-mail message...", "Info" , 0, "Info")
    Start-Sleep 5

    $Global:emailaddress = $textbox1Email.Text

    # Email will be sent to requester

    $sendTo = $Global:emailaddress
    $sendSubject = "Certificate request // $Global:TicketNumer"
    $sendCC = "<CC E-mail address>"
   
    Start-Sleep 5

    Send-MailMessage -Subject "$sendSubject" -From "DoNotReplyWintel@yourdomain.com" -To "$sendTo" -Cc $sendCC -SmtpServer "<your SMTP Server>" -Attachments C:\Temp\$Global:TicketNumer.zip -Body "Good day,

    Please see attached requested certificate // $Global:TicketNumer

    Should you have any questions, please contact us directly at <your custom support e-mail>.

    Best Regards,
    Wintel Team"

    [System.Windows.Forms.MessageBox]::Show("Workflow completed. `n`rE-mail was sent.", "Info" , 0, "Info")
    }
    }

    Catch {

    [System.Windows.Forms.MessageBox]::Show("E-mail was not sent. `n`rPlease contact your administrator.", "Error" , 0, "Error")
    }

})
[void]$form.Controls.Add($button1SendCert)

[void]$form.ShowDialog()
[void]$stream.Dispose()
[void]$Form.Dispose()
