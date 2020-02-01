#Constants
$ScriptFile = "$PSScriptRoot\pwn-Patch.ps1"
$SettingsFile = "$PSScriptRoot\Config.xml"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$main                            = New-Object system.Windows.Forms.Form
$main.ClientSize                 = '730,800'
$main.text                       = "pwn-Patch Config"
$main.TopMost                    = $false

$l_ScriptFile                    = New-Object system.Windows.Forms.Label
$l_ScriptFile.text               = "Script:"
$l_ScriptFile.AutoSize           = $true
$l_ScriptFile.width              = 25
$l_ScriptFile.height             = 10
$l_ScriptFile.location           = New-Object System.Drawing.Point(5,15)
$l_ScriptFile.Font               = 'Microsoft Sans Serif,10'

$t_ScriptFile                    = New-Object system.Windows.Forms.TextBox
$t_ScriptFile.Text               = $ScriptFile
$t_ScriptFile.multiline          = $false
$t_ScriptFile.Enabled            = $false
$t_ScriptFile.width              = 570
$t_ScriptFile.height             = 20
$t_ScriptFile.Anchor             = 'top,right,left'
$t_ScriptFile.location           = New-Object System.Drawing.Point(80,10)
$t_ScriptFile.Font               = 'Microsoft Sans Serif,10'

$l_ScriptFile_Status             = New-Object system.Windows.Forms.Label
$l_ScriptFile_Status.text        = "[not found]"
$l_ScriptFile_Status.AutoSize    = $true
$l_ScriptFile_Status.Visible     = $false
$l_ScriptFile_Status.width       = 25
$l_ScriptFile_Status.height      = 10
$l_ScriptFile_Status.Anchor      = 'top,right'
$l_ScriptFile_Status.location    = New-Object System.Drawing.Point(660,15)
$l_ScriptFile_Status.Font        = 'Microsoft Sans Serif,10'
$l_ScriptFile_Status.ForeColor   = "#ff0000"

$l_SettingsFile                    = New-Object system.Windows.Forms.Label
$l_SettingsFile.text               = "Settings:"
$l_SettingsFile.AutoSize           = $true
$l_SettingsFile.width              = 25
$l_SettingsFile.height             = 10
$l_SettingsFile.location           = New-Object System.Drawing.Point(5,40)
$l_SettingsFile.Font               = 'Microsoft Sans Serif,10'

$t_SettingsFile                    = New-Object system.Windows.Forms.TextBox
$t_SettingsFile.Text               = "$PSScriptRoot\Settings.xml"
$t_SettingsFile.multiline          = $false
$t_SettingsFile.Enabled            = $false
$t_SettingsFile.width              = 570
$t_SettingsFile.height             = 20
$t_SettingsFile.Anchor             = 'top,right,left'
$t_SettingsFile.location           = New-Object System.Drawing.Point(80,35)
$t_SettingsFile.Font               = 'Microsoft Sans Serif,10'

$l_SettingsFile_Status             = New-Object system.Windows.Forms.Label
$l_SettingsFile_Status.text        = "[not found]"
$l_SettingsFile_Status.AutoSize    = $true
$l_SettingsFile_Status.Visible     = $false
$l_SettingsFile_Status.width       = 35
$l_SettingsFile_Status.height      = 10
$l_SettingsFile_Status.Anchor      = 'top,right'
$l_SettingsFile_Status.location    = New-Object System.Drawing.Point(660,40)
$l_SettingsFile_Status.Font        = 'Microsoft Sans Serif,10'
$l_SettingsFile_Status.ForeColor   = "#ff0000"

$g_Notification                  = New-Object system.Windows.Forms.Groupbox
$g_Notification.height           = 250
$g_Notification.width            = 710
$g_Notification.Anchor           = 'top,right,left'
$g_Notification.text             = "Notification"
$g_Notification.location         = New-Object System.Drawing.Point(10,300)

$l_NotificationThreshold                 = New-Object system.Windows.Forms.Label
$l_NotificationThreshold.text            = "Notification Threshold:"
$l_NotificationThreshold.AutoSize        = $true
$l_NotificationThreshold.width           = 25
$l_NotificationThreshold.height          = 10
$l_NotificationThreshold.location        = New-Object System.Drawing.Point(15,55)
$l_NotificationThreshold.Font            = 'Microsoft Sans Serif,10'

$c_NotificationThreshold                 = New-Object system.Windows.Forms.ComboBox
$c_NotificationThreshold.text            = "disable"
$c_NotificationThreshold.width           = 100
$c_NotificationThreshold.height          = 20
$c_NotificationThreshold.location        = New-Object System.Drawing.Point(180,50)
$c_NotificationThreshold.Font            = 'Microsoft Sans Serif,10'

$l_NotificationThreshold_Result          = New-Object system.Windows.Forms.Label
$l_NotificationThreshold_Result.text     = "Never"
$l_NotificationThreshold_Result.AutoSize = $false
$l_NotificationThreshold_Result.width    = 395
$l_NotificationThreshold_Result.height   = 10
$l_NotificationThreshold_Result.Anchor   = 'top,right,left'
$l_NotificationThreshold_Result.location  = New-Object System.Drawing.Point(300,55)
$l_NotificationThreshold_Result.Font     = 'Microsoft Sans Serif,10'

$l_NotificationFromAddress               = New-Object system.Windows.Forms.Label
$l_NotificationFromAddress.text          = "From Address:"
$l_NotificationFromAddress.AutoSize      = $true
$l_NotificationFromAddress.width         = 25
$l_NotificationFromAddress.height        = 10
$l_NotificationFromAddress.location      = New-Object System.Drawing.Point(15,75)
$l_NotificationFromAddress.Font          = 'Microsoft Sans Serif,10'

$t_NotificationFromAddress               = New-Object system.Windows.Forms.TextBox
$t_NotificationFromAddress.multiline     = $false
$t_NotificationFromAddress.text          = "noreply@domain.com"
$t_NotificationFromAddress.width         = 520
$t_NotificationFromAddress.height        = 20
$t_NotificationFromAddress.Anchor        = 'top,right,left'
$t_NotificationFromAddress.location      = New-Object System.Drawing.Point(180,70)
$t_NotificationFromAddress.Font          = 'Microsoft Sans Serif,10'

$l_NotificationSubject                   = New-Object system.Windows.Forms.Label
$l_NotificationSubject.text              = "Subject:"
$l_NotificationSubject.AutoSize          = $true
$l_NotificationSubject.width             = 25
$l_NotificationSubject.height            = 10
$l_NotificationSubject.location          = New-Object System.Drawing.Point(15,115)
$l_NotificationSubject.Font              = 'Microsoft Sans Serif,10'

$t_NotificationSubject                   = New-Object system.Windows.Forms.TextBox
$t_NotificationSubject.multiline         = $false
$t_NotificationSubject.text              = "Your password is too weak.  Please change it."
$t_NotificationSubject.width             = 520
$t_NotificationSubject.height            = 20
$t_NotificationSubject.Anchor            = 'top,right,left'
$t_NotificationSubject.location          = New-Object System.Drawing.Point(180,110)
$t_NotificationSubject.Font              = 'Microsoft Sans Serif,10'

$l_NotificationEmailBody                 = New-Object system.Windows.Forms.Label
$l_NotificationEmailBody.text            = "Message Body:"
$l_NotificationEmailBody.AutoSize        = $true
$l_NotificationEmailBody.width           = 25
$l_NotificationEmailBody.height          = 10
$l_NotificationEmailBody.location        = New-Object System.Drawing.Point(14,135)
$l_NotificationEmailBody.Font            = 'Microsoft Sans Serif,10'

$t_NotificationEmailBody                 = New-Object system.Windows.Forms.TextBox
$t_NotificationEmailBody.multiline       = $true
$t_NotificationEmailBody.width           = 375
$t_NotificationEmailBody.height          = 100
$t_NotificationEmailBody.Anchor          = 'top,right,left'
$t_NotificationEmailBody.location        = New-Object System.Drawing.Point(180,130)
$t_NotificationEmailBody.Font            = 'Microsoft Sans Serif,10'

$l_NotificationLockAccount               = New-Object system.Windows.Forms.Label
$l_NotificationLockAccount.text          = "Expire password after:"
$l_NotificationLockAccount.AutoSize      = $true
$l_NotificationLockAccount.width         = 25
$l_NotificationLockAccount.height        = 10
$l_NotificationLockAccount.location      = New-Object System.Drawing.Point(15,230)
$l_NotificationLockAccount.Font          = 'Microsoft Sans Serif,10'

$c_NotificationLockAccount               = New-Object system.Windows.Forms.ComboBox
$c_NotificationLockAccount.text          = "disable"
$c_NotificationLockAccount.width         = 100
$c_NotificationLockAccount.height        = 20
$c_NotificationLockAccount.location      = New-Object System.Drawing.Point(180,230)
$c_NotificationLockAccount.Font          = 'Microsoft Sans Serif,10'

$r_NotificationUser              = New-Object system.Windows.Forms.RadioButton
$r_NotificationUser.text         = "User"
$r_NotificationUser.AutoSize     = $true
$r_NotificationUser.width        = 104
$r_NotificationUser.height       = 20
$r_NotificationUser.location     = New-Object System.Drawing.Point(55,15)
$r_NotificationUser.Font         = 'Microsoft Sans Serif,10,style=Bold'

$r_NotificationManager           = New-Object system.Windows.Forms.RadioButton
$r_NotificationManager.text      = "Manager"
$r_NotificationManager.AutoSize  = $true
$r_NotificationManager.width     = 104
$r_NotificationManager.height    = 20
$r_NotificationManager.location  = New-Object System.Drawing.Point(120,15)
$r_NotificationManager.Font      = 'Microsoft Sans Serif,10,style=Bold'

$r_NotificationAdmin                = New-Object system.Windows.Forms.RadioButton
$r_NotificationAdmin.text           = "Admin"
$r_NotificationAdmin.AutoSize       = $true
$r_NotificationAdmin.width          = 104
$r_NotificationAdmin.height         = 20
$r_NotificationAdmin.location       = New-Object System.Drawing.Point(206,14)
$r_NotificationAdmin.Font           = 'Microsoft Sans Serif,10,style=Bold'

$l_NotificationToAddress            = New-Object system.Windows.Forms.Label
$l_NotificationToAddress.text       = "To Address:"
$l_NotificationToAddress.AutoSize   = $true
$l_NotificationToAddress.width      = 25
$l_NotificationToAddress.height     = 10
$l_NotificationToAddress.location   = New-Object System.Drawing.Point(15,95)
$l_NotificationToAddress.Font       = 'Microsoft Sans Serif,10'

$t_NotificationToAddress                 = New-Object system.Windows.Forms.TextBox
$t_NotificationToAddress.multiline       = $false
$t_NotificationToAddress.width           = 520
$t_NotificationToAddress.height          = 20
$t_NotificationToAddress.Anchor          = 'top,right,left'
$t_NotificationToAddress.location        = New-Object System.Drawing.Point(180,90)
$t_NotificationToAddress.Font            = 'Microsoft Sans Serif,10'

$g_MailServer                       = New-Object system.Windows.Forms.Groupbox
$g_MailServer.height                = 40
$g_MailServer.width                 = 710
$g_MailServer.text                  = "Mail Server"
$g_MailServer.location              = New-Object System.Drawing.Point(10,230)

$l_MailAddress                   = New-Object system.Windows.Forms.Label
$l_MailAddress.text              = "Address:"
$l_MailAddress.AutoSize          = $true
$l_MailAddress.width             = 25
$l_MailAddress.height            = 10
$l_MailAddress.location          = New-Object System.Drawing.Point(19,20)
$l_MailAddress.Font              = 'Microsoft Sans Serif,10'

$t_MailAddress                        = New-Object system.Windows.Forms.TextBox
$t_MailAddress.multiline              = $false
$t_MailAddress.width                  = 300
$t_MailAddress.height                 = 20
$t_MailAddress.location               = New-Object System.Drawing.Point(80,15)
$t_MailAddress.Font                   = 'Microsoft Sans Serif,10'

$l_MailPort                      = New-Object system.Windows.Forms.Label
$l_MailPort.text                 = "Port:"
$l_MailPort.AutoSize             = $true
$l_MailPort.width                = 25
$l_MailPort.height               = 10
$l_MailPort.location             = New-Object System.Drawing.Point(410,15)
$l_MailPort.Font                 = 'Microsoft Sans Serif,10'

$t_MailPort                        = New-Object system.Windows.Forms.TextBox
$t_MailPort.multiline              = $false
$t_MailPort.width                  = 140
$t_MailPort.height                 = 20
$t_MailPort.location               = New-Object System.Drawing.Point(450,14)
$t_MailPort.Font                   = 'Microsoft Sans Serif,10'

$c_MailSSL                       = New-Object system.Windows.Forms.CheckBox
$c_MailSSL.text                  = "Use SSL"
$c_MailSSL.AutoSize              = $false
$c_MailSSL.width                 = 65
$c_MailSSL.height                = 10
$c_MailSSL.location              = New-Object System.Drawing.Point(620,20)
$c_MailSSL.Font                  = 'Microsoft Sans Serif,10'

$main.controls.AddRange(@($l_ScriptFile,$t_ScriptFile,$l_SettingsFile,$l_ScriptFile_Status,$t_SettingsFile,$l_SettingsFile_Status,$g_Notification,$g_MailServer))
$g_Notification.controls.AddRange(@($l_NotificationThreshold,$c_NotificationThreshold,$l_NotificationThreshold_Result,$l_NotificationFromAddress,$t_NotificationFromAddress,$l_NotificationSubject,$t_NotificationSubject,$l_NotificationEmailBody,$t_NotificationEmailBody,$l_NotificationLockAccount,$c_NotificationLockAccount,$r_NotificationUser,$r_NotificationManager,$r_NotificationAdmin,$l_NotificationToAddress,$t_NotificationToAddress))
$g_MailServer.controls.AddRange(@($l_MailAddress,$t_MailAddress,$l_MailPort,$t_MailPort,$c_MailSSL))




#Write your logic code here

[void]$main.ShowDialog()