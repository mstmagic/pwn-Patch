clear

# ****Constants****
$ScriptFile = "$PSScriptRoot\pwn-Patch.ps1"
$SettingsFile = "$PSScriptRoot\Config.xml"
$FontLarge = "Microsoft Sans Serif,10"
$FontSmall = "Microsoft Sans Serif,7"


# ****Load/Create XML Configuration****
$isSettingsFileMissing=(-not (Test-Path -LiteralPath $SettingsFile))
if (-not $isSettingsFileMissing) {
    [xml]$Config = Get-Content $SettingsFile
} else { 
    [xml]$Config = New-Object System.XML.XMLDocument
}
# ****Create Missing XML Tree Elements****
if ($Config.Config -eq $null)                                  { $Config.AppendChild($Config.CreateElement("Config")) | Out-Null }
if ($Config.Config.MailServer -eq $null)                       { $Config.SelectNodes("Config").AppendChild($Config.CreateElement("MailServer")) | Out-Null }
if ($Config.Config.Notification -eq $null)                     { $Config.SelectNodes("Config").AppendChild($Config.CreateElement("Notification")) | Out-Null }
if ($Config.Config.Notification.User -eq $null)                { $Config.SelectNodes("Config/Notification").AppendChild($Config.CreateElement("User")) | Out-Null }
if ($Config.Config.Notification.Admin -eq $null)               { $Config.SelectNodes("Config/Notification").AppendChild($Config.CreateElement("Admin")) | Out-Null }

# ****Set Missing XML Values to Default Values****
if ($Config.Config.SearchBase -eq $null)                       { $Config.SelectNodes("Config").SetAttribute(                    "SearchBase",     "OU=Organizational Unit,DC=Domain,DC=com") }
if ($Config.Config.DictionaryFile -eq $null)                   { $Config.SelectNodes("Config").SetAttribute(                    "DictionaryFile", "") }
if ($Config.Config.IndexFile -eq $null)                        { $Config.SelectNodes("Config").SetAttribute(                    "IndexFile",      "$PSScriptRoot\Config.xml") }
if ($Config.Config.EventsFile -eq $null)                       { $Config.SelectNodes("Config").SetAttribute(                    "EventFile",      "$PSScriptRoot\Failures.csv") }
if ($Config.Config.MailServer.Address -eq $null)               { $Config.SelectNodes("Config/MailServer").SetAttribute(         "Address",        "smtp.company.com") }
if ($Config.Config.MailServer.Port -eq $null)                  { $Config.SelectNodes("Config/MailServer").SetAttribute(         "Port",           "25") }
if ($Config.Config.MailServer.SSL -eq $null)                   { $Config.SelectNodes("Config/MailServer").SetAttribute(         "SSL",            "false") }
if ($Config.Config.Notification.User.Enabled -eq $null)        { $Config.SelectNodes("Config/Notification/User").SetAttribute(  "Threshold",      "1") }
if ($Config.Config.Notification.User.Threshold -eq $null)      { $Config.SelectNodes("Config/Notification/User").SetAttribute(  "Threshold",      "1") }
if ($Config.Config.Notification.User.ExpirePassword -eq $null) { $Config.SelectNodes("Config/Notification/User").SetAttribute(  "ExpirePassword", "true") }
if ($Config.Config.Notification.User.ExpireHours -eq $null)    { $Config.SelectNodes("Config/Notification/User").SetAttribute(  "ExpireHours",    "24") }
if ($Config.Config.Notification.User.FromAddress -eq $null)    { $Config.SelectNodes("Config/Notification/User").SetAttribute(  "FromAddress",    "itsupport@company.com") }
<# This Parameter is not needed and is ignored #>                $Config.SelectNodes("Config/Notification/User").SetAttribute(  "ToAddress",      "[End User]") 
if ($Config.Config.Notification.User.Subject -eq $null)        { $Config.SelectNodes("Config/Notification/User").SetAttribute(  "Subject",        "[Password Stength Monitor] Action Required: Your password is too weak") }
if ($Config.Config.Notification.User.MsgBodyFile -eq $null)    { $Config.SelectNodes("Config/Notification/User").SetAttribute(  "MsgBodyFile",    "$PSScriptRoot\AdminMessage.html") }
if ($Config.Config.Notification.Admin.Threshold -eq $null)     { $Config.SelectNodes("Config/Notification/Admin").SetAttribute( "Threshold",      "1") }
<# This Parameter is not needed and is ignored #>                $Config.SelectNodes("Config/Notification/Admin").SetAttribute( "ExpirePassword", "false")
<# This Parameter is not needed and is ignored #>                $Config.SelectNodes("Config/Notification/Admin").SetAttribute( "ExpireHours",    "-")
if ($Config.Config.Notification.Admin.FromAddress -eq $null)   { $Config.SelectNodes("Config/Notification/Admin").SetAttribute( "FromAddress",    "noreply@company.com") }
if ($Config.Config.Notification.Admin.ToAddress -eq $null)     { $Config.SelectNodes("Config/Notification/Admin").SetAttribute( "ToAddress",      "itsupport@company.com") }
if ($Config.Config.Notification.Admin.Subject -eq $null)       { $Config.SelectNodes("Config/Notification/Admin").SetAttribute( "Subject",        "[Password Stength Monitor] An end user has been found using a weak password") }
if ($Config.Config.Notification.Admin.MsgBodyFile -eq $null)   { $Config.SelectNodes("Config/Notification/Admin").SetAttribute( "MsgBodyFile",    "$PSScriptRoot\AdminMessage.html") }

# ****Default Message Body Blocks****
$DefaultUserMessage = @"
    <font style="font-size: 14pt">Your password is weak or common and you are required to change it.</font><br>
    <font style="font-size: 10pt"><br>
    <i><b><u>IMPORTANT:</u> The IT department will never ask you for your password.</b></i><br>
    <br>
    Hello [displayname],
    The IT department periodically scans your account to ensure you are using a strong password. You can be assured that IT does not know your password, but we do know that your password is among a list of weak, common, or exploited passwords.  We compared your password with a list of common user passwords on the internet and discovered that your password was listed [occurrences] time(s).<br>
    <br>
    We kindly request you change your password.  Your password will be set to expire: <b>[expiration]</b><br>
    <br>
    Please consider these guidelines when selecting a password to create a strong password:
    <ul>
    <li>Use passphrase with a few words, instead of simple passwords
    <li>Use as much of the keyboard as possible
    <li>Do not use easily guessed words, such as &ldquo;password&rdquo;, &ldquo;123456&rdquo;
    <li>Do not use keyboard patterns, such as &ldquo;qwerty&rdquo;, &ldquo;asdzxc&rdquo;, &ldquo;123456&rdquo;
    <li>Do not use names, including family names, company names, street names
    <li>Do not use private information such as SIN, Birth date, Anniversary, Phone Numbers
    <li>Do not use passwords that you have used on internet websites
    </ul>
    <br>
    <i>If you have more questions about this process, please contact IT support.</i><br></font>
"@

$DefaultAdminMessage=@"
    <font style="font-size: 14pt">One or more of your end users has been found using a weak password.</font><br>
    <font style="font-size: 10pt"><br>
    <b>[userlist]</i></b>
    <br>
    The IT department periodically scans employee's account to ensure they are using a strong passwords by comparing the users password against a list of known breached passwords provided by haveibeenpwned.com.  The users password has never been sent anywhere and there is no record of the users password.<br>
"@

If (Test-Path -LiteralPath $script:Config.Config.Notification.User.MsgBodyFile -ErrorAction Continue) {
    $UserMessage = Get-Content -LiteralPath $script:Config.Config.Notification.User.MsgBodyFile
} else {
    $UserMessage = $DefaultUserMessage
}
If (Test-Path -LiteralPath $script:Config.Config.Notification.Admin.MsgBodyFile -ErrorAction Continue) {
    $AdminMessage = Get-Content -LiteralPath $script:Config.Config.Notification.Admin.MsgBodyFile
} else {
    $AdminMessage = $DefaultAdminMessage
}


function isAdmin() {
    $isAdmin = (new-object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole("Administrators")
    return $isAdmin
}


# ****Create GUI****
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$main                            = New-Object system.Windows.Forms.Form
$main.ClientSize                 = '730,600'
$main.Text                       = "pwn-Patch Config"
$main.TopMost                    = $false
$l_ScriptFile                    = New-Object system.Windows.Forms.Label
$l_ScriptFile.Text               = "Script:"
$l_ScriptFile.AutoSize           = $true
$l_ScriptFile.Font               = $FontLarge
$t_ScriptFile                    = New-Object system.Windows.Forms.TextBox
$t_ScriptFile.Text               = $ScriptFile
$t_ScriptFile.AutoSize           = $false
$t_ScriptFile.Multiline          = $false
$t_ScriptFile.Enabled            = $false
$t_ScriptFile.Anchor             = 'top,right,left'
$t_ScriptFile.Font               = $FontLarge
$l_ScriptFile_NotFound             = New-Object system.Windows.Forms.Label
$l_ScriptFile_NotFound.Text        = "[not found]"
$l_ScriptFile_NotFound.AutoSize    = $true
$l_ScriptFile_NotFound.Visible     = $false
$l_ScriptFile_NotFound.Anchor      = 'top,right'
$l_ScriptFile_NotFound.Font        = $FontSmall
$l_ScriptFile_NotFound.ForeColor   = "#ff0000"
$l_SettingsFile                    = New-Object system.Windows.Forms.Label
$l_SettingsFile.Text               = "Settings:"
$l_SettingsFile.AutoSize           = $true
$l_SettingsFile.Font               = $FontLarge
$t_SettingsFile                    = New-Object system.Windows.Forms.TextBox
$t_SettingsFile.Text               = $SettingsFile
$t_SettingsFile.AutoSize           = $false
$t_SettingsFile.Multiline          = $false
$t_SettingsFile.Enabled            = $false
$t_SettingsFile.Anchor             = 'top,right,left'
$t_SettingsFile.Font               = $FontLarge
$l_SettingsFile_NotFound             = New-Object system.Windows.Forms.Label
$l_SettingsFile_NotFound.Text        = "[not found]"
$l_SettingsFile_NotFound.AutoSize    = $true
$l_SettingsFile_NotFound.Visible     = $false
$l_SettingsFile_NotFound.Anchor      = 'top,right'
$l_SettingsFile_NotFound.Font        = $FontSmall
$l_SettingsFile_NotFound.ForeColor   = "#ff0000"
$b_DSInternals_Help                  = New-Object System.Windows.Forms.Button
$b_DSInternals_Help.Text             = "?"
$b_DSInternals_Help.BackColor        = [System.Drawing.Color]::Blue
#$b_DSInternals_Help.AutoSizeMode     = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$b_DSInternals_Help.ForeColor        = [System.Drawing.Color]::White

$b_DSInternals_Help.AutoSize         = $false
$b_DSInternals_Help.Width            = 30
$b_DSInternals_Help.Font             = $FontSmall
$l_DSInternals                    = New-Object system.Windows.Forms.Label
$l_DSInternals.Text               = "DSInternals:"
$l_DSInternals.AutoSize           = $true
$l_DSInternals.Font               = $FontLarge
$b_DSInternals_Install            = New-Object System.Windows.Forms.Button
$b_DSInternals_Install.Text       = "Install"
$b_DSInternals_Install.AutoSize   = $true
$b_DSInternals_Install.Font       = $FontSmall
$l_DSInternals_Status             = New-Object system.Windows.Forms.Label
$l_DSInternals_Status.AutoSize    = $true
$l_DSInternals_Status.Font        = $FontSmall
$l_SearchBase                    = New-Object system.Windows.Forms.Label
$l_SearchBase.Text               = "Search Base:"
$l_SearchBase.AutoSize           = $true
$l_SearchBase.Font               = $FontLarge
$t_SearchBase                    = New-Object system.Windows.Forms.TextBox
$t_SearchBase.AutoSize           = $False
$t_SearchBase.Multiline          = $false
$t_SearchBase.Enabled            = $true
$t_SearchBase.Text               = $Config.Config.SearchBase
$t_SearchBase.Anchor             = 'top,right,left'
$t_SearchBase.Font               = $FontLarge
$l_Dictionary                    = New-Object system.Windows.Forms.Label
$l_Dictionary.Text               = "Password Dictionary:"
$l_Dictionary.AutoSize           = $true
$l_Dictionary.Font               = $FontLarge
$t_Dictionary                    = New-Object system.Windows.Forms.TextBox
$t_Dictionary.AutoSize           = $False
$t_Dictionary.Multiline          = $false
$t_Dictionary.Text               = $Config.Config.DictionaryFile
$t_Dictionary.Anchor             = 'top,right,left'
$t_Dictionary.Font               = $FontLarge
$l_Dictionary_Missing            = New-Object system.Windows.Forms.Label
$l_Dictionary_Missing.Text       = "[missing]"
$l_Dictionary_Missing.AutoSize   = $true
$l_Dictionary_Missing.Font       = $FontSmall
$l_Dictionary_Missing.Anchor      = 'top,right'
$l_Dictionary_Missing.ForeColor  = "#ff0000"
$b_Dictionary_Browse             = New-Object System.Windows.Forms.Button
$b_Dictionary_Browse.Text        = "Browse"
$b_Dictionary_Browse.AutoSize    = $true
$b_Dictionary_Browse.Anchor      = 'top,right'
$b_Dictionary_Browse.Font        = $FontSmall
$l_Index                    = New-Object system.Windows.Forms.Label
$l_Index.Text               = "Index File:"
$l_Index.AutoSize           = $true
$l_Index.Font               = $FontLarge
$t_Index                    = New-Object system.Windows.Forms.TextBox
$t_Index.AutoSize           = $False
$t_Index.Multiline          = $false
$t_Index.Text               = $Config.Config.IndexFile
$t_Index.Anchor             = 'top,right,left'
$t_Index.Font               = $FontLarge
$l_Index_Missing            = New-Object system.Windows.Forms.Label
$l_Index_Missing.Text       = "[missing]"
$l_Index_Missing.AutoSize   = $true
$l_Index_Missing.Font       = $FontSmall
$l_Index_Missing.Anchor      = 'top,right'
$l_Index_Missing.ForeColor  = "#ff0000"
$b_Index_Browse             = New-Object System.Windows.Forms.Button
$b_Index_Browse.Text        = "Browse"
$b_Index_Browse.AutoSize    = $true
$b_Index_Browse.Anchor      = 'top,right'
$b_Index_Browse.Font        = $FontSmall
$b_Index_ReIndex             = New-Object System.Windows.Forms.Button
$b_Index_ReIndex.Text        = "Re-Index"
$b_Index_ReIndex.AutoSize    = $true
$b_Index_ReIndex.Anchor      = 'top,right'
$b_Index_ReIndex.Font        = $FontSmall
$l_Events                    = New-Object system.Windows.Forms.Label
$l_Events.Text               = "Events File:"
$l_Events.AutoSize           = $true
$l_Events.Font               = $FontLarge
$t_Events                    = New-Object system.Windows.Forms.TextBox
$t_Events.AutoSize           = $False
$t_Events.Multiline          = $false
$t_Events.Text               = $Config.Config.EventFile
$t_Events.Anchor             = 'top,right,left'
$t_Events.Font               = $FontLarge
$b_Events_Browse             = New-Object System.Windows.Forms.Button
$b_Events_Browse.Text        = "Browse"
$b_Events_Browse.AutoSize    = $true
$b_Events_Browse.Anchor      = 'top,right'
$b_Events_Browse.Font        = $FontSmall
$l_MailAddress               = New-Object system.Windows.Forms.Label
$l_MailAddress.Text          = "Mail Server Address:"
$l_MailAddress.AutoSize      = $true
$l_MailAddress.Font          = $FontLarge
$t_MailAddress               = New-Object system.Windows.Forms.TextBox
$t_MailAddress.AutoSize      = $false
$t_MailAddress.Anchor        = 'left,top,right'
$t_MailAddress.Text          = $Config.Config.MailServer.Address
$t_MailAddress.Multiline     = $false
$t_MailAddress.Font          = $FontLarge
$l_MailPort                  = New-Object system.Windows.Forms.Label
$l_MailPort.Text             = "Port:"
$l_MailPort.AutoSize         = $true
$l_MailPort.Anchor           = 'top,right'
$l_MailPort.Font             = $FontLarge
$t_MailPort                  = New-Object system.Windows.Forms.TextBox
$t_MailPort.AutoSize         = $false
$t_MailPort.Anchor           = 'top,right'
$t_MailPort.Text             = $Config.Config.MailServer.Port
$t_MailPort.Multiline        = $true
$t_MailPort.Font             = $FontLarge
$c_MailSSL                   = New-Object system.Windows.Forms.CheckBox
$c_MailSSL.Text              = "Use SSL"
$c_MailSSL.AutoSize          = $true
$c_MailSSL.Checked           = ($config.Config.MailServer.SSL -like "true")
$c_MailSSL.Anchor            = 'top,right'
$c_MailSSL.Font              = $FontLarge
$g_Notification              = New-Object system.Windows.Forms.Groupbox
$g_Notification.Anchor       = 'top,bottom,right,left'
$g_Notification.Text         = "Notification"
$r_NotificationUser              = New-Object system.Windows.Forms.RadioButton
$r_NotificationUser.Text         = "User"
$r_NotificationUser.AutoSize     = $true
$r_NotificationUser.Checked      = $true
$r_NotificationUser.Font         = 'Microsoft Sans Serif,10,style=Bold'
$r_NotificationAdmin                = New-Object system.Windows.Forms.RadioButton
$r_NotificationAdmin.Text           = "Admin"
$r_NotificationAdmin.AutoSize       = $true
$r_NotificationAdmin.Font           = 'Microsoft Sans Serif,10,style=Bold'
$c_NotificationEnabled                   = New-Object system.Windows.Forms.CheckBox
$c_NotificationEnabled.Text              = "Enable Notifiations"
$c_NotificationEnabled.AutoSize          = $true
$c_NotificationEnabled.Checked           = ($config.Config.MailServer.SSL -like "true")
$c_NotificationEnabled.Anchor            = 'top,right'
$c_NotificationEnabled.Font              = $FontLarge
$l_NotificationThreshold                 = New-Object system.Windows.Forms.Label
$l_NotificationThreshold.Text            = "Notification Threshold:"
$l_NotificationThreshold.AutoSize        = $true
$l_NotificationThreshold.Font            = $FontLarge
$c_NotificationThreshold                 = New-Object system.Windows.Forms.ComboBox
$c_NotificationThreshold.AutoSize        = $true
$c_NotificationThreshold.Font            = $FontLarge
$c_NotificationThreshold.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$c_NotificationThreshold.Items.Add("0 (Disabled)") | Out-Null
$c_NotificationThreshold.Items.Add("1 (Most Forgiving)") | Out-Null
$c_NotificationThreshold.Items.Add("2") | Out-Null
$c_NotificationThreshold.Items.Add("3") | Out-Null
$c_NotificationThreshold.Items.Add("4") | Out-Null
$c_NotificationThreshold.Items.Add("5") | Out-Null
$c_NotificationThreshold.Items.Add("6") | Out-Null
$c_NotificationThreshold.Items.Add("7") | Out-Null
$c_NotificationThreshold.Items.Add("8") | Out-Null
$c_NotificationThreshold.Items.Add("9") | Out-Null
$c_NotificationThreshold.Items.Add("10 (Most Aggressive)") | Out-Null
$c_NotificationThreshold.SelectedItem="0 (Disabled)"
$l_NotificationLockAccount               = New-Object system.Windows.Forms.Label
$l_NotificationLockAccount.Text          = "Expire password after:"
$l_NotificationLockAccount.AutoSize      = $true
$l_NotificationLockAccount.Font          = $FontLarge
$t_NotificationLockAccount               = New-Object system.Windows.Forms.TextBox
$t_NotificationLockAccount.Multiline     = $false
$t_NotificationLockAccount.Text          = "0"
$t_NotificationLockAccount.AutoSize      = $false
$t_NotificationLockAccount.Font          = $FontLarge
$l_NotificationLockAccount_Hours               = New-Object system.Windows.Forms.Label
$l_NotificationLockAccount_Hours.Text          = "hours later"
$l_NotificationLockAccount_Hours.AutoSize      = $true
$l_NotificationLockAccount_Hours.Font          = $FontLarge
$c_NotificationLockAccount_Disabled           = New-Object system.Windows.Forms.CheckBox
$c_NotificationLockAccount_Disabled.Text      = "never"
$c_NotificationLockAccount_Disabled.AutoSize  = $true
$c_NotificationLockAccount_Disabled.Anchor    = 'top,left'
$c_NotificationLockAccount_Disabled.Font      = $FontLarge
$l_NotificationThreshold_Result          = New-Object system.Windows.Forms.Label
$l_NotificationThreshold_Result.Text     = "Never"
$l_NotificationThreshold_Result.AutoSize = $true
$l_NotificationThreshold_Result.Font     = $FontLarge
$l_NotificationFromAddress               = New-Object system.Windows.Forms.Label
$l_NotificationFromAddress.Text          = "From Address:"
$l_NotificationFromAddress.AutoSize      = $true
$l_NotificationFromAddress.Font          = $FontLarge
$t_NotificationFromAddress               = New-Object system.Windows.Forms.TextBox
$t_NotificationFromAddress.Multiline     = $false
$t_NotificationFromAddress.Text          = "noreply@company.com"
$t_NotificationFromAddress.AutoSize      = $false
$t_NotificationFromAddress.Anchor        = 'top,right,left'
$t_NotificationFromAddress.Font          = $FontLarge
$l_NotificationToAddress                 = New-Object system.Windows.Forms.Label
$l_NotificationToAddress.Text            = "To Address:"
$l_NotificationToAddress.AutoSize        = $true
$l_NotificationToAddress.Font            = $FontLarge
$l_NotificationToAddress.Enabled         = $false
$t_NotificationToAddress                 = New-Object system.Windows.Forms.TextBox
$t_NotificationToAddress.Multiline       = $false
$t_NotificationToAddress.Text            = "[End User]"
$t_NotificationToAddress.AutoSize        = $false
$t_NotificationToAddress.Anchor          = 'top,right,left'
$t_NotificationToAddress.Font            = $FontLarge
$t_NotificationToAddress.Enabled         = $false
$l_NotificationSubject                   = New-Object system.Windows.Forms.Label
$l_NotificationSubject.Text              = "Subject:"
$l_NotificationSubject.AutoSize          = $true
$l_NotificationSubject.Font              = $FontLarge
$t_NotificationSubject                   = New-Object system.Windows.Forms.TextBox
$t_NotificationSubject.Multiline         = $false
$t_NotificationSubject.Text              = "ACTION REQUIRED: Your password is too weak.  Please change it."
$t_NotificationSubject.AutoSize          = $false
$t_NotificationSubject.Anchor            = 'top,right,left'
$t_NotificationSubject.Font              = $FontLarge
$l_NotificationEmailBody                 = New-Object system.Windows.Forms.Label
$l_NotificationEmailBody.Text            = "Message Body:"
$l_NotificationEmailBody.AutoSize        = $true
$l_NotificationEmailBody.Font            = $FontLarge
$t_NotificationEmailBody_File                 = New-Object system.Windows.Forms.TextBox
$t_NotificationEmailBody_File.Multiline       = $false
$t_NotificationEmailBody_File.AutoSize        = $false
$t_NotificationEmailBody_File.Anchor          = 'top,right,left,bottom'
$t_NotificationEmailBody_File.Font            = $FontLarge
$t_NotificationEmailBody_File.Text           = $Config.Config.Notification.User.MsgBodyFile
$l_NotificationEmailBody_Missing            = New-Object system.Windows.Forms.Label
$l_NotificationEmailBody_Missing.Text       = "[missing]"
$l_NotificationEmailBody_Missing.AutoSize   = $true
$l_NotificationEmailBody_Missing.Font       = $FontSmall
$l_NotificationEmailBody_Missing.Anchor      = 'top,right'
$l_NotificationEmailBody_Missing.ForeColor  = "#ff0000"
$b_NotificationEmailBody_Browse             = New-Object System.Windows.Forms.Button
$b_NotificationEmailBody_Browse.Text        = "Browse"
$b_NotificationEmailBody_Browse.AutoSize    = $true
$b_NotificationEmailBody_Browse.Anchor      = 'top,right'
$b_NotificationEmailBody_Save             = New-Object System.Windows.Forms.Button
$b_NotificationEmailBody_Save.Text        = "Save"
$b_NotificationEmailBody_Save.AutoSize    = $true
$b_NotificationEmailBody_Save.Anchor      = 'top,right'
$l_NotificationEmailBody_Missing.Font        = $FontSmall
$t_NotificationEmailBody_Text                 = New-Object system.Windows.Forms.TextBox
$t_NotificationEmailBody_Text.Multiline       = $true
$t_NotificationEmailBody_Text.AutoSize        = $false
$t_NotificationEmailBody_Text.Anchor          = 'top,right,left,bottom'
$t_NotificationEmailBody_Text.Font            = $FontLarge
$t_NotificationEmailBody_Text.Text            = $UserMessage
$b_SaveConfig             = New-Object System.Windows.Forms.Button
$b_SaveConfig.Text        = "Save Settings"
$b_SaveConfig.AutoSize    = $true
$b_SaveConfig.Anchor      = 'bottom,right'
$b_SaveConfig.Font        = $FontLarge
$main.controls.AddRange(@($l_ScriptFile,$t_ScriptFile,$l_ScriptFile_NotFound,$l_SettingsFile,$t_SettingsFile,$l_SettingsFile_NotFound,$b_DSInternals_Help,$l_DSInternals,$l_DSInternals_Status,$b_DSInternals_Install,$l_SearchBase,$t_SearchBase,$l_Dictionary,$t_Dictionary,$l_Dictionary_Missing,$b_Dictionary_Browse,$l_Index,$t_Index,$l_Index_Missing,$b_Index_Browse,$b_Index_ReIndex,$l_Events,$t_Events,$b_Events_Browse,$l_MailAddress,$t_MailAddress,$l_MailPort,$t_MailPort,$c_MailSSL,$g_Notification,$b_SaveConfig))
$g_Notification.controls.AddRange(@($r_NotificationUser,$r_NotificationAdmin,$l_NotificationThreshold,$c_NotificationThreshold,$l_NotificationThreshold_Result,$l_NotificationLockAccount,$t_NotificationLockAccount,$l_NotificationLockAccount_Hours,$c_NotificationLockAccount_Disabled,$l_NotificationFromAddress,$t_NotificationFromAddress,$l_NotificationToAddress,$t_NotificationToAddress,$l_NotificationSubject,$t_NotificationSubject,$l_NotificationEmailBody,$t_NotificationEmailBody_File,$l_NotificationEmailBody_Missing,$b_NotificationEmailBody_Browse,$b_NotificationEmailBody_Save,$t_NotificationEmailBody_Text))


Function ResizeGUI () {
    $l_ScriptFile.Top = 5
    $l_ScriptFile.Left = 5
    $l_SettingsFile.Top = $l_ScriptFile.Bottom + 5
    $l_SettingsFile.Left = $l_ScriptFile.Left
    $t_ScriptFile.Top = $l_ScriptFile.Top
    $t_ScriptFile.Left = $l_SettingsFile.Right
    $t_ScriptFile.Height = $l_ScriptFile.Height
    $ScriptFileMissing=(-not (Test-Path -LiteralPath $script:ScriptFile))
    if ($ScriptFileMissing) {
        $l_ScriptFile_NotFound.Visible = $true
        $l_ScriptFile_NotFound.Left = $main.ClientSize.Width - $l_ScriptFile_NotFound.Width - 5
        $l_ScriptFile_NotFound.Top = $l_ScriptFile.Top + (($l_ScriptFile.Height - $l_ScriptFile_NotFound.Height)/2)
        $t_ScriptFile.Width = $l_ScriptFile_NotFound.Left - $t_ScriptFile.Left - 5
    } else {
        $l_ScriptFile_NotFound.Visible
        $t_ScriptFile.Width = $main.ClientSize.Width - $t_ScriptFile.Left - 5
    }
    $t_SettingsFile.Top = $l_SettingsFile.Top
    $t_SettingsFile.Left = $t_ScriptFile.Left
    $t_SettingsFile.Height = $l_SettingsFile.Height
    $SettingsFileMissing=(-not (Test-Path -LiteralPath $script:SettingsFile))
    if ($SettingsFileMissing) {
        $l_SettingsFile_NotFound.Visible = $true
        $l_SettingsFile_NotFound.Left = $main.ClientSize.Width - $l_SettingsFile_NotFound.Width - 5
        $l_SettingsFile_NotFound.Top = $l_SettingsFile.Top + (($l_SettingsFile.Height - $l_SettingsFile_NotFound.Height)/2)
        $t_SettingsFile.Width = $l_SettingsFile_NotFound.Left - $t_SettingsFile.Left - 5
    } else {
        $l_SettingsFile_NotFound.Visible
        $t_SettingsFile.Width = $main.ClientSize.Width - $t_SettingsFile.Left - 5
    }
    #Move DSInternals Help Button - Align to left of the form underneath the Settings
        $b_DSInternals_Help.Top = $l_SettingsFile.Bottom + 5
        $b_DSInternals_Help.Left = $l_SettingsFile.Left
    #Move DSInternals Label - Align to the right of the DSInternals Help Button
        $l_DSInternals.Top = $b_DSInternals_Help.Top + (($b_DSInternals_Help.Height - $l_DSInternals.Height) / 2)
        $l_DSInternals.Left = $b_DSInternals_Help.Right + 5
    #Move DSInternals Status Align to the right of the DSInternals Label
        $l_DSInternals_Status.Top = $b_DSInternals_Help.Top + (($b_DSInternals_Help.Height - $l_DSInternals_Status.Height) / 2) 
        $l_DSInternals_Status.Left = $l_DSInternals.Right + 5
    #Color and Recaption the DSInternals Status
    #If DSInternals is installed it should be green and read "[Installed]", otherwise it should be red and read "[Missing]"
        if (Get-InstalledModule -Name "DSInternals" -MinimumVersion 4.0 -ErrorAction SilentlyContinue) {
            $l_DSInternals_Status.Text="[Installed]"
            $l_DSInternals_Status.ForeColor="#006600"
            $b_DSInternals_Install.Enabled = $false
        } else {
            $l_DSInternals_Status.Text="[Missing]"
            $l_DSInternals_Status.ForeColor="#ff0000"
            $b_DSInternals_Install.Enabled = $true
        }
    #Move DSInternals Install Button.  Align to the right of the DSInternals Status Label
        $b_DSInternals_Install.Top = $b_DSInternals_Help.Top + (($b_DSInternals_Help.Height - $b_DSInternals_Install.Height) / 2) 
        $b_DSInternals_Install.Left = $l_DSInternals_Status.Right + 5
    $l_Dictionary.Top = $l_DSInternals.Bottom + 5
    $l_Dictionary.Left = $l_DSInternals.Left
    $b_Dictionary_Browse.Top = $l_Dictionary.Top + (($l_Dictionary.Height - $b_Dictionary_Browse.Height) / 2) 
    $b_Dictionary_Browse.Left = $main.ClientSize.Width - $b_Dictionary_Browse.Width - $b_Index_ReIndex.Width - 10
    $t_Dictionary.Top = $l_Dictionary.Top
    $t_Dictionary.Left = $l_Dictionary.Right + 5
    $t_Dictionary.Height = $l_Dictionary.Height
    if ($script:Config.Config.DictionaryFile) {
        $DictionaryFileExist = Test-Path -LiteralPath $script:Config.Config.DictionaryFile -ErrorAction SilentlyContinue
    } else {
        $DictionaryFileExist = $false
    }
    if ($DictionaryFileExist) {
        $l_Dictionary_Missing.Visible = $false
        $t_Dictionary.Width = $b_Dictionary_Browse.Left - $t_Dictionary.Left - 5
    } else {
        $l_Dictionary_Missing.Visible = $true
        $l_Dictionary_Missing.Top = $l_Dictionary.Top + (($l_Dictionary.Height - $l_Dictionary_Missing.Height) / 2)
        $l_Dictionary_Missing.Left = $b_Dictionary_Browse.Left - $l_Dictionary_Missing.Width - 5
        $t_Dictionary.Width = $l_Dictionary_Missing.Left - $t_Dictionary.Left - 5
    }
    $l_Index.Top = $l_Dictionary.Bottom + 5
    $l_Index.Left = $l_Dictionary.Left
    $b_Index_ReIndex.Top = $l_Index.Top + (($l_Index.Height - $b_Index_ReIndex.Height) / 2) 
    $b_Index_ReIndex.Left = $main.ClientSize.Width - $b_Index_ReIndex.Width - 5
    $b_Index_Browse.Top = $b_Index_ReIndex.Top 
    $b_Index_Browse.Left = $b_Dictionary_Browse.Left
    $t_Index.Top = $l_Index.Top
    $t_Index.Left = $t_Dictionary.Left
    $t_Index.Height = $l_Index.Height
    if ($script:Config.Config.IndexFile -eq "") {
        $b_Index_ReIndex.Enabled = $false
    } else {
        $b_Index_ReIndex.Enabled = $true 
    }
    $IndexFileExist = Test-Path -LiteralPath $script:Config.Config.IndexFile -ErrorAction Continue
    if ($IndexFileExist) {
        $l_Index_Missing.Visible = $false
        $t_Index.Width = $b_Index_Browse.Left - $t_Index.Left - 5
    } else {
        $l_Index_Missing.Visible = $true
        $l_Index_Missing.Top = $l_Index.Top + (($l_Index.Height - $l_Index_Missing.Height) / 2)
        $l_Index_Missing.Left = $b_Index_Browse.Left - $l_Index_Missing.Width - 5
        $t_Index.Width = $l_Index_Missing.Left - $t_Index.Left - 5
    }
    $l_Events.Top = $l_Index.Bottom + 5
    $l_Events.Left = $l_Index.Left
    $b_Events_Browse.Top = $l_Events.Top + (($l_Events.Height - $b_Events_Browse.Height) / 2) 
    $b_Events_Browse.Left = $main.ClientSize.Width - $b_Events_Browse.Width - $b_Index_ReIndex.Width - 10
    $t_Events.Top = $l_Events.Top
    $t_Events.Left = $t_Index.Left
    $t_Events.Height = $l_Events.Height
    $t_Events.Width = $b_Events_Browse.Left - $t_Events.Left - 5
    $l_SearchBase.Top = $l_Events.Bottom + 5
    $l_SearchBase.Left = $l_Events.Left
    $t_SearchBase.Top = $l_SearchBase.Top
    $t_SearchBase.Left = $t_Dictionary.Left
    $t_SearchBase.Height = $l_SearchBase.Height
    $t_SearchBase.Width = $main.ClientSize.Width - $t_SearchBase.Left - 5
    $l_MailAddress.Top = $l_SearchBase.Bottom + 5
    $l_MailAddress.Left = $l_SearchBase.Left
    $c_MailSSL.Top = $l_MailAddress.Top + (($l_MailAddress.Height - $c_MailSSL.Height) / 2) 
    $c_MailSSL.Left = $main.ClientSize.Width - $c_MailSSL.Width - 5
    $t_MailPort.Top = $l_MailAddress.Top
    $t_MailPort.Width = 100
    $t_MailPort.Left = $c_MailSSL.Left - $t_MailPort.Width - 5
    $t_MailPort.Height = $l_MailPort.Height
    $l_MailPort.Top = $l_MailAddress.Top
    $l_MailPort.Left = $t_MailPort.Left - $l_MailPort.Width - 5
    $t_MailAddress.Top = $l_MailAddress.Top
    $t_MailAddress.Left = $l_MailAddress.Right + 5
    $t_MailAddress.Height = $l_MailAddress.Height
    $t_MailAddress.Width = $l_MailPort.Left - $t_MailAddress.Left - 5
    $b_SaveConfig.Top = $main.ClientSize.Height - $b_SaveConfig.Height - 5
    $b_SaveConfig.Left = $main.ClientSize.Width - $b_SaveConfig.Width - 5
    $g_Notification.Top = $l_MailAddress.Bottom + 5
    $g_Notification.Left = 5
    $g_Notification.Width = $main.ClientSize.Width - 10
    $g_Notification.Height = $b_SaveConfig.Top - $g_Notification.Top - 5
    $r_NotificationUser.Top = 15
    $r_NotificationUser.Left = 15
    $r_NotificationAdmin.Top = $r_NotificationUser.Top
    $r_NotificationAdmin.Left = $r_NotificationUser.Right + 5
    $l_NotificationThreshold.Top = $r_NotificationUser.Bottom + 5
    $l_NotificationThreshold.Left = $r_NotificationUser.Left
    $c_NotificationThreshold.Top = $l_NotificationThreshold.Top + (($l_NotificationThreshold.Height - $c_NotificationThreshold.Height) / 2)
    $c_NotificationThreshold.Left = $l_NotificationThreshold.Right + 5
    $c_NotificationThreshold.Height = $l_NotificationThreshold.Height
    $c_NotificationThreshold.Width = $l_NotificationThreshold.Width
    $l_NotificationThreshold_Result.Top = $l_NotificationThreshold.Top
    $l_NotificationThreshold_Result.Left = $c_NotificationThreshold.Right + 5
    $l_NotificationLockAccount.Top = $l_NotificationThreshold.Bottom + 5
    $l_NotificationLockAccount.Left = $l_NotificationThreshold.Left
    $t_NotificationLockAccount.Top = $l_NotificationLockAccount.Top
    $t_NotificationLockAccount.Left = $l_NotificationLockAccount.Right + 5
    $t_NotificationLockAccount.Height = $l_NotificationLockAccount.Height
    $t_NotificationLockAccount.Width = $c_NotificationThreshold.Width
    $l_NotificationLockAccount_Hours.Top = $l_NotificationLockAccount.Top
    $l_NotificationLockAccount_Hours.Left = $t_NotificationLockAccount.Right + 5
    $c_NotificationLockAccount_Disabled.Top = $l_NotificationLockAccount.Top
    $c_NotificationLockAccount_Disabled.Left = $l_NotificationLockAccount_Hours.Right + 5
    $l_NotificationFromAddress.Top = $l_NotificationLockAccount.Bottom + 5
    $l_NotificationFromAddress.Left = $l_NotificationThreshold.Left
    $t_NotificationFromAddress.Top = $l_NotificationFromAddress.Top
    $t_NotificationFromAddress.Height = $l_NotificationFromAddress.Height
    $t_NotificationFromAddress.Left = $l_NotificationFromAddress.Right + 5
    $t_NotificationFromAddress.Width = $g_Notification.ClientSize.Width - $t_NotificationFromAddress.Left - 15
    $l_NotificationToAddress.Top = $l_NotificationFromAddress.Bottom + 5
    $l_NotificationToAddress.Left = $l_NotificationFromAddress.Left
    $t_NotificationToAddress.Top = $l_NotificationToAddress.Top
    $t_NotificationToAddress.Height = $l_NotificationToAddress.Height
    $t_NotificationToAddress.Left = $t_NotificationFromAddress.Left
    $t_NotificationToAddress.Width = $g_Notification.ClientSize.Width - $t_NotificationToAddress.Left - 15
    $l_NotificationSubject.Top = $l_NotificationToAddress.Bottom + 5
    $l_NotificationSubject.Left = $l_NotificationToAddress.Left
    $t_NotificationSubject.Top = $l_NotificationSubject.Top
    $t_NotificationSubject.Height = $l_NotificationSubject.Height
    $t_NotificationSubject.Left = $t_NotificationFromAddress.Left
    $t_NotificationSubject.Width = $g_Notification.ClientSize.Width - $t_NotificationSubject.Left - 15
    $l_NotificationEmailBody.Top = $t_NotificationSubject.Bottom + 5
    $l_NotificationEmailBody.Left = $l_NotificationSubject.Left
    $b_NotificationEmailBody_Save.Top = $l_NotificationEmailBody.Top  + (($l_NotificationEmailBody.Height - $l_NotificationEmailBody.Height) / 2)
    $b_NotificationEmailBody_Save.Left = $g_Notification.ClientSize.Width - $b_NotificationEmailBody_Save.Width - 15
    $b_NotificationEmailBody_Browse.Top = $b_NotificationEmailBody_Save.Top
    $b_NotificationEmailBody_Browse.Left = $b_NotificationEmailBody_Save.Left - $b_NotificationEmailBody_Browse.Width - 5
    $t_NotificationEmailBody_File.Top = $l_NotificationEmailBody.Top
    $t_NotificationEmailBody_File.Left = $t_NotificationFromAddress.Left
    $t_NotificationEmailBody_File.Width = $g_Notification.ClientSize.Width - $t_NotificationEmailBody_File.Left - 15
    $t_NotificationEmailBody_File.Height = $l_NotificationEmailBody.Height
    if ($r_NotificationUser.Checked) {
        $MsgFileExists = Test-Path -LiteralPath $script:Config.Config.Notification.User.MsgBodyFile -ErrorAction Continue
    } else {
        $MsgFileExists = Test-Path -LiteralPath $script:Config.Config.Notification.Admin.MsgBodyFile -ErrorAction Continue
    }
    if ($MsgFileExists) {
        $l_NotificationEmailBody_Missing.Visible = $false
        $t_NotificationEmailBody_File.Width = $b_NotificationEmailBody_Browse.Left - $t_NotificationEmailBody_File.Left - 5
    } else {
        $l_NotificationEmailBody_Missing.Visible = $true
        $l_NotificationEmailBody_Missing.Top = $l_NotificationEmailBody.Top + (($l_NotificationEmailBody.Height - $l_NotificationEmailBody_Missing.Height) / 2)
        $l_NotificationEmailBody_Missing.Left = $b_NotificationEmailBody_Browse.Left - $l_NotificationEmailBody_Missing.Width - 5
        $t_NotificationEmailBody_File.Width = $l_NotificationEmailBody_Missing.Left - $t_NotificationEmailBody_File.Left - 5
    }
    $t_NotificationEmailBody_Text.Top = $l_NotificationEmailBody.Bottom + 5
    $t_NotificationEmailBody_Text.Left = $l_NotificationEmailBody.Left
    $t_NotificationEmailBody_Text.Width = $g_Notification.ClientSize.Width - $t_NotificationEmailBody_Text.Left - 15
    $t_NotificationEmailBody_Text.Height = $g_Notification.ClientSize.Height - $t_NotificationEmailBody_Text.Top - 15
}

$b_DSInternals_Help.Add_Click(
    {
        [System.Windows.MessageBox]::Show("DSInternals is used to retirieve NTLM password hashes from Active Directory and needs to be installed.`n`nhttps://www.dsinternals.com/en/retrieving-active-directory-passwords-remotely/","DSInternals Help",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information)
    }
)

ResizeGUI


#Write your logic code here

[void]$main.ShowDialog()
$main.BringToFront()