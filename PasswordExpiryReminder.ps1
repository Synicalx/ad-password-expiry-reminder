# Email settings
$smartHost         = '<Your SMTP Host>'
$sendFrom          = '<Your FROM email address>'
$mailTitle         = 'Password Expiry Reminder'

# The OU to look in with Get-Aduser
$searchOU          = "<OU=Users,DC=acme,DC=local>"

# Get the current date
$date              = Get-Date

# Date to check against, ie the date threshold
$dateThreshold

# Specify log file
$logFile           = "expiryLog.log"

# Write an entry to the log file, with a date stamp
function LogWrite
{
    Param ([string]$newLogMessage)
    $funcTime = Get-Date -format o
    Add-Content $logFile -value "$funcTime :: $newLogMessage"
}

# If today is Friday, check for users expiring Monday
switch ($date.DayofWeek)
{
    'Friday'        {
                        $dateThreshold = (Get-Date).AddDays(+3) 
                        LogWrite " Starting scipt, Friday mode + 3 days"
                        LogWrite " Threshold date is $dateThreshold"
                    }
    # Don't email on weekends
    'Saturday'      { Exit }
    'Sunday'        { Exit }

    default         {
                        $dateThreshold = (Get-Date).AddDays(+1) 
                        LogWrite " Starting scipt, normal mode + 1 days"
                        LogWrite " Threshold date is $dateThreshold"
                    }
}

# Generate a list of users and convert timestamp
try 
{
    $users         = Get-ADUser -searchbase $searchOU `
                        -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} `
                        -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed", "mail"|
                        where msDS-UserPasswordExpiryTimeComputed -gt 0 |
                        Select-Object Name,@{ Name="Expiry";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed") } }, `
                        mail, givenname, surname, DisplayName
} catch {
    LogWrite "Error in getting AD users :: $_.Exception.ItemName :: $_.Exception.Message"
}

# List of users to email due to meeting threshold, convert to array later
$usersToEmail      = New-Object System.Collections.Generic.List[System.Object]

# Refine Expiry to a DayofYear for easy comparison
foreach ($user in $users)
{
    $userDate      = $user.Expiry

    if ($userDate.DayofYear -eq $dateThreshold.DayofYear)
    {
        $usersToEmail.Add($user)
    }
}

# Convert system list to array
$usersToEmail.ToArray()


# Compose the email as HTML - this can also be stored in a seperate html file
$mailBody =

"  <html>
    <head>
      <meta name='viewport' content='width=device-width'>
      <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>
      <title>New User</title>
      <style>
      /* -------------------------------------
          INLINED WITH htmlemail.io/inline
      ------------------------------------- */
      /* -------------------------------------
          RESPONSIVE AND MOBILE FRIENDLY STYLES
      ------------------------------------- */
      @media only screen and (max-width: 620px) {
        table[class=body] h1 {
          font-size: 28px !important;
          margin-bottom: 10px !important;
        }
        table[class=body] p,
              table[class=body] ul,
              table[class=body] ol,
              table[class=body] td,
              table[class=body] span,
              table[class=body] a {
          font-size: 16px !important;
        }
        table[class=body] .wrapper,
              table[class=body] .article {
          padding: 10px !important;
        }
        table[class=body] .content {
          padding: 0 !important;
        }
        table[class=body] .container {
          padding: 0 !important;
          width: 100% !important;
        }
        table[class=body] .main {
          border-left-width: 0 !important;
          border-radius: 0 !important;
          border-right-width: 0 !important;
        }
        table[class=body] .btn table {
          width: 100% !important;
        }
        table[class=body] .btn a {
          width: 100% !important;
        }
        table[class=body] .img-responsive {
          height: auto !important;
          max-width: 100% !important;
          width: auto !important;
        }
      }
      /* -------------------------------------
          PRESERVE THESE STYLES IN THE HEAD
      ------------------------------------- */
      @media all {
        .ExternalClass {
          width: 100%;
        }
        .ExternalClass,
              .ExternalClass p,
              .ExternalClass span,
              .ExternalClass font,
              .ExternalClass td,
              .ExternalClass div {
          line-height: 100%;
        }
        .apple-link a {
          color: inherit !important;
          font-family: inherit !important;
          font-size: inherit !important;
          font-weight: inherit !important;
          line-height: inherit !important;
          text-decoration: none !important;
        }
        .btn-primary table td:hover {
          background-color: #34495e !important;
        }
        .btn-primary a:hover {
          background-color: #34495e !important;
          border-color: #34495e !important;
        }
      }

      .new {  
          color: #333;
          font-family: Helvetica, Arial, sans-serif;
          width: 640px; 
          border-collapse: 
          collapse; border-spacing: 0; 
          margin-left: 8px;
      }

      td.new, th.new {  
          border: 1px solid transparent; /* No more visible border */
          height: 30px; 
          transition: all 0.3s;  /* Simple transition for hover effect */
      }

      th.new {  
          background: #DFDFDF;  /* Darken header a bit */
          font-weight: bold;
      }

      td.new {  
          text-align: left;
      }


      </style>
    </head>
    <body class='' style='background-color: #f6f6f6; font-family: sans-serif; -webkit-font-smoothing: antialiased; font-size: 14px; line-height: 1.4; margin: 0; padding: 0; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%;'>


      <table border='0' cellpadding='0' cellspacing='0' class='body' style='border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%; background-color: #f6f6f6;'>
        <tr>
          <td style='font-family: sans-serif; font-size: 14px; vertical-align: top;'>&nbsp;</td>
          <td class='container' style='font-family: sans-serif; font-size: 14px; vertical-align: top; display: block; Margin: 0 auto; max-width: 580px; padding: 10px; width: 580px;'>
            <div class='content' style='box-sizing: border-box; display: block; Margin: 0 auto; max-width: 580px; padding: 10px;'>

              <!-- START CENTERED WHITE CONTAINER -->

              <span class='preheader' style='color: transparent; display: none; height: 0; max-height: 0; max-width: 0; opacity: 0; overflow: hidden; mso-hide: all; visibility: hidden; width: 0;'></span>
              <table class='main' style='border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%; background: #ffffff; border-radius: 3px;'>

                <!-- START MAIN CONTENT AREA -->
                <tr>
                  <td class='wrapper' style='font-family: sans-serif; font-size: 14px; vertical-align: top; box-sizing: border-box; padding: 20px;'>
                    <table border='0' cellpadding='0' cellspacing='0' style='border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;'>
                      <tr>
                        <td style='font-family: sans-serif; font-size: 14px; vertical-align: top;'>
                          <h1>Your password is expiring soon</h1>
                            <table class='new' border='0' cellpadding='0' cellspacing='0' style=' border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%; '>
                              <tr cclass='new' style='font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; Margin-bottom: 0px;'>
                                <td class='new'>Please be advised that your account password is set to expire within one business day. To prevent account lockout, please change your password at your earliest convenience.</td> 
                              </tr>
                            </table>
                          <p style='font-family: sans-serif; font-size: 14px; font-weight: normal; margin: 0; Margin-bottom: 15px;'><br><br>Kind Regards, <br>Service Desk</p>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>

              <!-- END MAIN CONTENT AREA -->
              </table>

              <!-- START FOOTER -->
              <div class='footer' style='clear: both; Margin-top: 10px; text-align: center; width: 100%;'>
                <table border='0' cellpadding='0' cellspacing='0' style='border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;'>
                  <tr>
                    <td class='content-block' style='font-family: sans-serif; vertical-align: top; padding-bottom: 10px; padding-top: 10px; font-size: 12px; color: #999999; text-align: center;'>
                      <span class='apple-link' style='color: #999999; font-size: 12px; text-align: center;'>If you experience any difficulties changing your password, please contact the Service Desk<a></span>
                    </td>
                  </tr>
                </table>
              </div>
              <!-- END FOOTER -->

            <!-- END CENTERED WHITE CONTAINER -->
            </div>
          </td>
          <td style='font-family: sans-serif; font-size: 14px; vertical-align: top;'>&nbsp;</td>
        </tr>
      </table>
    </body>
  </html>
  "

# Iterate over all users
if ($usersToEmail)
{
    foreach ($expiringUser in $usersToEmail)
    {
        try 
        {
            Send-MailMessage -Subject $mailTitle -From $sendFrom -To $expiringUser.mail -SmtpServer $smartHost -body $mailBody -BodyAsHtml -Priority high
            LogWrite " Emailing notification to $($expiringUser.name)"

        } catch {
            LogWrite " Error in sending an email :: $_.Exception.ItemName :: $_.Exception.Message"
        }

    }

} else {
    LogWrite " No users expiring"
}

LogWrite " Run completed"
