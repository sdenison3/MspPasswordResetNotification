#region - Import Modules
$modules = @("Microsoft.Graph.Authentication","Microsoft.Graph.Mail","Microsoft.Graph.Users.Actions","MSAL.PS","ActiveDirectory") #Import any modules that could be needed
$loadedmodules = Get-Module
foreach ($module in $modules) {
    if ($loadedmodules.name -notcontains $module) {
        try {
            Import-Module $module -Force
        }
        catch {
            Install-Module $module -Force
            Import-Module $module -Force
        }
    }
}
#endregion#>

#region - Get Date and Set Amount of days to warn
$Today = Get-Date
$warnDays = 7 # How many days remaining - User's less than or equal will be emailed
#endregion#>


#region - Get a list of AD accounts where enables and the password can expire
$ADUsers = Get-ADUser -Filter {Enabled -eq $true -and PasswordNeverExpires -eq $false} -Properties 'msDS-UserPasswordExpiryTimeComputed', 'mail'
$AlreadyExpiredList = "" #initializes a list of already expired users to be added to later
#endregion#>

#region - User Message logic and email
    # For each account
foreach( $User in $ADUsers ){
    
    # Get the expiry date and convert to date time
    $ExpireDate = [datetime]::FromFileTime( $User.'msDS-UserPasswordExpiryTimeComputed' )
    $ExpireDate_String = $ExpireDate.ToString("MM/dd/yyyy h:mm tt")
    
    # Calculate the days remaining
    $daysRmmaining = New-TimeSpan -Start $Today -End $ExpireDate
    $daysRmmaining = $daysRmmaining.Days
 
    $usersName = $User.Name

    if ($daysRmmaining -le $warnDays -And $daysRmmaining -ge 0) #Logic can be changed here to send out on certain days
    {
            # Generate email subject from days remaining
        if ($daysRmmaining -eq 0)
        {
            $emailSubject = "Your password expires today"
        } else {
            $emailSubject = "Your password expires in $daysRmmaining days"
        }
 
        # Get users email
        if($null -eq $user.mail) #Noted below -  Might create a lot of tickets - need to decide what to do with this
        {
            # The user does not have an email address in AD, send alert
            #$sendTo = "something@something.com"
            #$htmlMessage = "$usersName password expires $ExpireDate_String. #But can't email them as their AD mail field is balnk :-("
        } else {
            # The user has an email address
            $sendTo = $user.mail
            
            $htmlMessage = @"

            <html>

            <h2>Hey $userName!</h2>
            <div>We are proactively reaching out to let you know that your Windows login password will be expiring in $daysRemaining. Please be sure to reset your password before it expires on $ExpireDate_String.</div>
            <div>
            <div>
            <h4>Here are the steps to change your Login password using your Windows workstation:</h4>
            <div>1) Login as usual.</div>
            <div>2) Press the CTRL + ALT + DELETE keys simultaneously.</div>
            <div>3) Click the “Change a password” option.</div>
            <div>4) In the labeled boxes:</div>
            <div>- Enter your current password.</div>
            <div>- Enter your new password.</div>
            <div>- Re-enter your new password.</div>
            <div>5) If you have your work email on a mobile device, you may also need to update your device with the new password.</div>
            <div> </div>
            <div>To change your password while on VPN: (Insert link to document here)</div>
            <div>Thank you for your co-operation. If you have any issues please call us at (XXX)XXX-XXXX or email us at something@something.com</div>
            <br />
            <div>-Your Friendly neighborhood Admin</div>
            </div>
            </div>

            </html>
"@
        }
#endregion#>
    }
    #region - expired password message
    elseif ($daysRmmaining -lt 0) 
    {
        $emailSubject = "Your password has expired"
        # Password already expired, add the users details to a list ready for email
        $userMail = $user.mail
        $AlreadyExpiredList = $AlreadyExpiredList + "$usersName, $userMail, $ExpireDate_String</br>"
        
        $sendTo = $user.mail

        $htmlMessage = @"
        <html>

        <h2>Hey $userName!</h2>
        <div>Your Windows login password has expired. Bummer! In order to get it reset, please call Pendello at your earliest convenience at (XXX)XXX-XXXX.</div>
        <br />
        <div>We will be awaiting your call!</div>
        <br />
        <div>Your friendly neighborhood Admin</div>

        </html>

"@ 
    }
    #endregion#>

    #region - Connect-Graph   
    $AppId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    $TenantId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    $ClientSecret = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
    $Scopes = 'https://graph.microsoft.com/.default'
    
    $MsalToken = Get-MsalToken -TenantId $TenantId -ClientId $AppId -Scopes $scopes -ClientSecret ($ClientSecret | ConvertTo-SecureString -AsPlainText -Force)
    
    Connect-Graph -AccessToken $MsalToken.AccessToken

    #region - Mail Properties
    #region - Sender
    $Sender1 = 'something@something.com'
    #endregion


    #region - Body
        $body  = @{
            content = $htmlmessage
            ContentType = 'html'
        }
    #endregion#>
    #endregion#>
    
    #region - Create and send message
    #Create Message - Puts it in Drafts
    $Message = New-MgUserMessage -UserId $Sender1 -Body $body -ToRecipients $sendTo -Subject $emailSubject #-Attachments $Attachment
    
    #Send Message
    Send-MgUserMessage -UserId $Sender1 -MessageId $Message.Id
    #endregion#>
}