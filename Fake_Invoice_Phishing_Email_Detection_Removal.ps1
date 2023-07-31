#The researcher utilizes a PowerShell module developed by Dr. Tobias Weltner, which grants access to the OCR (Optical Character Recognition) functionality built into Windows 10. Readers #find the module at the following link: #https://github.com/TobiasPSP/PsOcr.
  
#This module, called PsOcr, allows PowerShell to extract image text. 
#The PowerShell script also incorporates Danny Davis's instructions on downloading attachments from Outlook. YReaders can find the instructions at the following link: #https://www.danny-#davis.com/blog/2019/9/27/download-#attachements-#from-outlook-with-powershell.
#The PowerShell script successfully retrieves attachments from emails and saves them to a local directory. 
#The script leverages the Outlook application and MAPI namespace to access email data, load emails from the designated subfolder, iterate through each email and process the #attachments.

#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
#Open a remote Session to Exchange-Server

$password=ConvertTo-SecureString '[Password]' -AsPlainText -Force 
$credential = New-Object System.Management.Automation.PSCredential ('[exchangeadminusername]', $password) 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange-server/PowerShell/ -Authentication Kerberos -Credential $credential
Import-pssession $session -DisableNameChecking -AllowClobber

#Phase-1: Copy all emails with image attachments to a designated mailbox

#This Phase needs to be done in Exchange Server

#Phase-2: Extract attachments from the designated mailbox

# link to the folder 

$olFolderPath = "\\[EmailAddress]\Inbox\[Folder Name]"

# Set the location to a temporary file

$filePath = "//[FilePath]/"

# Use MAPI namespace

$outlook = new-object -com outlook.application; 
$mapi = $outlook.GetNameSpace("MAPI");

# set the Inbox folder id

$olDefaultFolderInbox = 6
$inbox = $mapi.GetDefaultFolder($olDefaultFolderInbox) 

# access the target subfolder

$olTargetFolder = $inbox.Folders | Where-Object { $_.FolderPath -eq $olFolderPath }

# load emails

$emails = $olTargetFolder.Items
$count = $emails.count

# process the emails

$count = 0
foreach ($email in $emails) {
    $email
    $count = $count + 1

    # Format the timestamp

    $timestamp = $email.ReceivedTime.ToString("yyyyMMddhhmmss")

    # Filter out the attachments

    $email.Attachments | foreach {

        # Insert the timestamp into the file name
        $fileName = $_.FileName
        $fileName = $fileName.Insert($fileName.IndexOf('.'), $timestamp)

        # Save the attachment

        $_.SaveAsFile((Join-Path '[File Path]' $fileName))
    }

    # Phase-3: Convert Text from Image file using PSOImageToText module and save it to a file

    foreach ($file in Get-ChildItem "[File Path]") {
        $OCRText = Convert-PsoImageToText -Path (Join-Path '[File Path]' $file)

        # Read the text file containing the text from the image file

        Write-Output $OCRText >> "[File Path]\ImageToText.txt"

        # Read Indicator of Compromise file (Make sure this file is updated frequently from threat intelligence feeds)

        $IOC = Get-Content -Path [File Path for the String Match File]\StringsToMatch.txt

        # Search IOCs from the ImageToText file

        foreach ($string in $IOC) {
            if ($OCRText -match $string) {
                Write-Host 'Contains String: ' $string

                # Phase-4: Delete emails

                $subject = $email.Subject
                $from = $email.SenderEmailAddress
                $SenderName = $email.SenderName
                $attachment = $email.Attachments | Select-Object -ExpandProperty DisplayName
                $Search = New-ComplianceSearch -Name "[NewComplianceSearchName]" -ExchangeLocation All -ContentMatchQuery "(attachment:$attachment)(from:$from)"
                start-ComplianceSearch -Identity $Search.Identity
                New-ComplianceSearchAction -SearchName "[NewComplianceSearchName]" -Purge

                # Clean up ComplianceSearch

                Remove-ComplianceSearch -Identity "[NewComplianceSearchName]" -Confirm:$false

                # Track Removed Fake-Invoice emails

                echo ----- $email.ReceivedTime "From: " $from "To: " $email.To $subject $attachment >> "[Path]\Trackers.txt"

                exit
            }
        }

        Remove-Item [File Path]\*.*
    }

    Echo "Purge Succeeded"
}