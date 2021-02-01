$script = ''
function Send-PowerMail {
<# 
.DESCRIPTION 
 Sends a Branded mail by means of a Microsoft Exchange Server
.EXAMPLE

 $SendEmail = @{
     Subject         = 'Testing'
     Signature       = 'Test Services'
     SendTo          = "testservices@test.com"
     MailFrom        = 'testservices@test.com'
     Content         = 'This is a test mail' 
     Attachment      = 'c:\temp\test.csv'
     ImageHeader     = 'https://test.com/test.png'
 }

 .\Send-PowerMail.ps1 @SendEmail

.PARAMETER Subject
 Subject of the email object
.PARAMETER Signature
 Signature text of the email. Fill in a name, team or department
.PARAMETER SendTo,
 Recipient(s) of the email object
.PARAMETER MailFrom
 Sender of the email object
.PARAMETER Content
 Content of the email object
.PARAMETER Attachment
 Attachment file(s) of the email object
.PARAMETER HeaderImage
 Image that can be used in the header of the email
.OUTPUTS
 Email object
#> 

[CmdLetBinding()]
Param(

    [Parameter (Mandatory = $True)]
    [String] $Subject,

    [Parameter (Mandatory = $True)]
    [String] $Signature,

    [Parameter (Mandatory = $True)]
    [String[]] $SendTo,

    [Parameter (Mandatory = $True)]
    [String] $MailFrom, 

    [Parameter (Mandatory = $True)]
    [Array] $Content,

    [Parameter (Mandatory = $False)]
    [ValidateScript({$_ | test-path })]
    [Array] $Attachment, 

    [Parameter (Mandatory = $True)]
    [ValidateScript({$_ | test-connection})]
    [String] $SMTPServer, 

    [Parameter (Mandatory = $false)]
    [String] $HeaderImage

)
#region Preparation
    
    $Body = @"
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
    <html>
        <head>
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Mail Notification &nbsp;</title>
            <style type="text/css">body { width:100% !important; height: 100% !important; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; margin: 0; padding: 0; background-color: #F2F2F2;} img { border: 0; outline: none; text-decoration: none; padding-bottom: 0; display: inline; -ms-interpolation-mode: bicubic; } table { border-collapse: collapse; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%; mso-table-lspace: 0pt; mso-table-rspace: 0pt;} table td { border-collapse: collapse; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%; mso-table-lspace: 0pt; mso-table-rspace: 0pt;} a { -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%; }</style>
        </head>
        <body><!-- IMAGE FILLING TABLE // -->
            <table border="0" cellpadding="0" cellspacing="0" style="height: 100%; width: 100%; align-self: center; margin: 0; padding: 0; background-color: #F2F2F2;">
                <tr>
                    <td style="height: 100%; width: 100%; vertical-align: top; align-self: center; margin: 0; padding: 18px; border-top: 0;">
                        <!-- CONTENT // -->
                            <center>
                                <table border="0" cellpadding="18" cellspacing="0" width="600" style="background-color: #FFFFFF; width: 600px !important;">
                                    <tr>
                                        <td style="vertical-align: top;">
                                            <img alt="" src=$HeaderImage width="564" style="max-width: 564px; vertical-align: bottom; align-self: center;">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="vertical-align: top; color: #606060; font-family: Helvetica; font-size: 14px; line-height: 150%; text-align: left !important;">Dear Administrator,<br><br>
                                        
                                        $Content
                                        
                                        <br>    
                                        <br>
                                        This is an Automated email. Please do not reply.
                                        
                                        <br>
                                        <br>
                                        Kind Regards,
                                        <br>$Signature 
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                                <tr>
                                                    <td width="439" style="text-align: right; height: 34px; vertical-align: bottom; color:#808080; font-size: 14px; font-family: Helvetica;">&nbsp;</td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </center>
                    </td>
                </tr>
            </table>
        </body>
    </html>
"@

#endregion

#region Execution  

$message = New-Object System.Net.Mail.MailMessage $MailFrom, $SendTo
foreach ($recipient in $SendTo) 
{
    $null = $message.To.Add(
        $recipient    
    )
}

if($Attachment)
{
    foreach($file in $Attachment)
    {
        $message.Attachments.Add($file)
    }
}

$message.Subject = $Subject
$message.IsBodyHTML = $true
$message.Body = $Body
$smtp = New-Object Net.Mail.SmtpClient($SMTPServer)

try{
    $smtp.Send($message)
    Write-Output "Sending Email"
}
catch{
    Write-Error $_ 
    Write-Output "Could not send email!"
}
#endregion

}
