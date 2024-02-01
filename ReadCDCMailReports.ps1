 <#
      .SYNOPSIS
      Read emails and outputs the subject, from, to, date, and body as a pdf
           
      $userid - Outlook user account (ie rhowell@cytovance.com) 
      $SQL_CDC_Reports_Id - id of SQL CDC Reports Outlook folder
                            See getmail.ps1 for get Inbox id and child folder id's
     
      .DESCRIPTION
      Read emails from provided user (ie rhowell@cytovance.com) 
      Prints data to a printer called "PrintPDFUnattended" which prints the
      data to a static file in the temp folder without user interaction. The generated
      PDF file is then moved to a user-specified location.
      If the printer "PrintPDFUnattended" is not present, the function calls
      the dependent function Install-PDFPrinter. Requires Windows 10, Server 2016, or better.
 
 
      .EXAMPLE
 
      .NOTES
      Uses WienExpertsLive module v1.14
      https://www.powershellgallery.com/packages/WienExpertsLive/1.14
      Install-Module -Name WienExpertsLive

  #>



import-module exchangeonlinemanagement
connect-mgGraph -Scopes "User.Read.All" -NoWelcome

$tab = "`t"
$cr ="`n"
$lf = "`r"

#  SQL CDC Reports folder ID
$SQL_CDC_Reports_Id = "AAMkAGVlYmU4NzRiLTAyYjktNDk1Yy1iZDUyLTMzMmM4Y2NjMDRmMQAuAAAAAAD_qA-5_p81Qod5sFO8_Mx8AQBm9W_Yi9IfTZ-3rKDo5rZ8AAAAGGZDAAA="
$userid = "rhowell@cytovance.com"

# Base path for pdf reports
$repbasepath = "Z:\IT Support\PI AF Audit Trail CDC Reports\" 
$pathExists = Test-Path -Path $repbasepath

If (-not ($pathExists)) 
{
    Write-host "Z drive does not exist!"
    Return
}

# Get top n emails for SQL CDC Reports Folder
$cdc_emails = Get-MgUserMailFolderMessage -MailFolderId $SQL_CDC_Reports_Id -UserId $userid -Top 10

# $mail
foreach ($email in $cdc_emails)
{
    #  Get the message
    $msg = Get-MgUserMessage -UserId $userid -MessageId $email.Id

    #  Get Month and Year report was emailed in MMMYY format
    $monthyear = ((Get-Culture).DateTimeFormat.GetMonthName($msg.SentDateTime.Month)).Substring(0,3).ToUpper() + ($msg.SentDateTime.Year.ToString()).Substring(2,2)
    #  Get day sent and pad with leading 0
    $day = $msg.SentDateTime.Day.ToString()
    $day = $day.PadLeft(2, "0")

    #  Change path to Month Year of report
    $path = $repbasepath + $monthyear + "\"
    $pathExists = Test-Path -Path $path

    If (-not ($pathExists)) 
    {
        Write-host $path " does not exist!"
        New-Item -ItemType "directory" -Path $path
        $pathExists = Test-Path -Path $path
        If (-not ($pathExists)) 
        {
            Write-host "Creation of path " $path " failed!"
            Write-host " "
            Return
        }
        Write-host "Created " $path
        Write-Host " "
    }

    # Set-Location -Path $path
    $file = $path + $day + $monthyear + ".*"
    #  ie 03JAN24.*
    #  search for the file
    $filefound = Get-ChildItem -Path $file -Name 

    if ($null -eq $filefound)
    {
        # File not found - print the file
        <#
            Email header when printed:
            Weekly SQL Server Status
            SVC_Alerts <SVC_Alerts@cytovance.com>
            Thu 1/4/2024 7:18 AM
            To: PI Admin <piadmin@cytovance.com>
        #>
        #  Use $tempoutput to output to c:\temp
        # $tempoutput = "c:\temp\" + $day + $monthyear + ".pdf"
        # $file points to the Z drive IT folder for reports
        # $file = "Z:\IT Support\PI AF Audit Trail CDC Reports\JAN24\08JAN24.*"
        $file = $file.Replace('*', 'pdf')
        
        $wholemessage = "Subject: " + $msg.Subject + $cr + $lf
        $wholemessage = $wholemessage + "From:    " + $msg.From.EmailAddress.Name + " <" + $msg.From.EmailAddress.Address + ">" + $cr + $lf
        $wholemessage = $wholemessage + "Date:    " + $msg.SentDateTime.ToLocalTime() + $cr + $lf
        $wholemessage = $wholemessage + "To:      " + $msg.ToRecipients.EmailAddress.Name + " <" + $msg.ToRecipients.EmailAddress.Address + ">" + $cr + $lf
        $wholemessage = $wholemessage + $msg.Body.Content  + $cr + $lf
        
        $wholemessage | Out-PDFFile -Path $file -Open

    }
    # $monthyear = $msg.SentDateTime.Month # + $msg.SentDateTime.Year.ToString()
    $msg.Subject + $tab + $msg.SentDateTime.ToLocalTime() + $tab + $monthyear + $tab + $filefound
}

