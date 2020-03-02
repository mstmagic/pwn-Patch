#********************************************************
#Parameters - Used to script setup if desired
#********************************************************
[CmdletBinding()]
Param(
        [Parameter()] [string]$SearchBase,
        [Parameter()] [string]$DictionaryFile,
        [Parameter()] [string]$IndexFile
    )

#********************************************************
#Constants
#********************************************************
$ScriptFile = "$PSScriptRoot\pwn-Patch.ps1"
$ConfigFile = "$PSScriptRoot\Config.xml"

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

clear

#********************************************************
#Check if DSInternals is installed.  If not - Install it!
#Need to ensure that user is elevated before installing
#********************************************************
Write-Host "DSINTERNALS" -BackgroundColor White -ForegroundColor Black
If ((Get-InstalledModule -Name "DSInternals" -MinimumVersion 4.0 -ErrorAction SilentlyContinue))  {
    write-host "DSInternals is allready installed" -ForegroundColor Green
} else {
    write-host "Missing DSInternals - Attempting to Install" -ForegroundColor Yellow
    $isAdmin = (new-object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole("Administrators")
    If (-not $isAdmin) {
        Write-Host "Error: Administrative Privilages Required" -ForegroundColor Red
        Write-Host "Please Run the script with Administrative Privilages to allow for the installation of DSInterals" -ForegroundColor Red
        Write-Host
        Write-Host "[EXITING]" -ForegroundColor Red
        return
    }
    Install-Module -Name DSInternals
    If ((Get-InstalledModule -Name "DSInternals" -MinimumVersion 4.0 -ErrorAction SilentlyContinue))  {
        write-host "[Installed]" -ForegroundColor Green
    } else {
        Write-Host "Unable to Install DSInternals due to error" -ForegroundColor Red
        Write-Host "Please correct the issue and try again" -ForegroundColor Red
        Write-Host
        Write-Host "[EXITING]" -ForegroundColor Red
        return
    }
}

#********************************************************
#Load the configuration file - If it exists
#           $SettingsFile = "$PSScriptRoot\Config.xml"
#It contains the location of the dictionary, index, and basic settings
#********************************************************
$isConfigFileMissing=(-not (Test-Path -Path $ConfigFile))
if (-not $isConfigFileMissing) {
    [xml]$Config = Get-Content $ConfigFile
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
if ($Config.Config.DictionaryFile -eq $null)                   { $Config.SelectNodes("Config").SetAttribute(                    "DictionaryFile", "$PSScriptRoot\pwned-passwords-ntlm-ordered-by-hash-v5.txt") }
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



#********************************************************
#Get the domain and OU of where the user accounts are
#Example: OU=Users,DC=Domain,DC=com
#********************************************************
Write-Host
Write-Host "SEARCH BASE" -BackgroundColor White -ForegroundColor Black
if (-not $SearchBase) {
    Write-Host "Please enter the location of your users in Active Directory" -ForegroundColor Yellow
    Write-Host "  Example: OU=Users,DC=Domain,DC=com" -ForegroundColor Yellow
    $SearchBase=Read-Host "Search Base [$($Config.Config.SearchBase)]"
}
if ($SearchBase -ne "") { $Config.Config.SearchBase = $SearchBase}
write-host "Selected Value: $($Config.Config.SearchBase)"


#********************************************************
#Get the location of the password dictionary
#Downloaded from: https://haveibeenpwned.com/Passwords
#********************************************************
Write-Host
Write-Host "DICTIONARY FILE" -BackgroundColor White -ForegroundColor Black
if (-not $DictionaryFile) {
    Write-Host "Please select the location of your Dictionary File" -ForegroundColor Yellow
    Write-Host "    This is downloaded from: https://haveibeenpwned.com/Passwords" -ForegroundColor Yellow
    Write-Host "    Format: NTLM Format and Orderd by hash" -ForegroundColor Yellow
    $OpenFileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        Title="Ordered NTLM Password Dictionary from https://haveibeenpwned.com/Passwords"
        Filter = 'Password Dictionary (*.txt)|*.txt|All Files (*.*)|*.*'
    }
    $Path = $Config.Config.DictionaryFile
    While (-not (Test-Path -Path $Path))  {
        $Path = Split-Path -Path $Path -Parent
        if (-not $Path) { $Path = $PSScriptRoot }
    }
    if ((Get-Item -Path $Path) -is [System.IO.DirectoryInfo]) {
        $OpenFileBrowser.InitialDirectory = $Path
        $OpenFileBrowser.FileName=""
    } else {
        $OpenFileBrowser.InitialDirectory = Split-Path -Path $Path -Parent
        $OpenFileBrowser.FileName=Split-Path -Path $Path -Leaf
    }
    if ($OpenFileBrowser.ShowDialog() -ne "OK") {
        Write-Host "User Cancled" -ForegroundColor Red
        Write-Host
        Write-Host "[EXITING]" -ForegroundColor Red
        return
    }
    $DictionaryFile = $OpenFileBrowser.FileName
} else {
    if (-not (Test-Path $DictionaryFile -ErrorAction SilentlyContinue)) {
        Write-Host "Dictionary File doesn't exist" -ForegroundColor Red
        Write-Host
        Write-Host "[EXITING]" -ForegroundColor Red
        return
    } else {
        Write-Host "Set to $DictionaryFile"
    }
}   
$config.Config.DictionaryFile = $DictionaryFile


#********************************************************
#Get the location of the index file
#The index file contains references to the location of hashes in the dictionary
#Required to speed up password searches
#********************************************************
Write-Host
Write-Host "INDEX FILE" -BackgroundColor White -ForegroundColor Black
if (-not $IndexFile) {
    Write-Host "Please select the location of your Password Dictionary Index File" -ForegroundColor Yellow
    Write-Host "    The index file is used to speed up searches in the password dictionary" -ForegroundColor Yellow
    $SaveFileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{ 
        Title="Password Dictionary Index File"
        Filter = 'Index File (*.csv)|*.csv|All Files (*.*)|*.*'
    }
    $Path = $Config.Config.IndexFile
    Do  {
        $Path = Split-Path -Path $Path -Parent
        if (-not $Path) { $Path = $PSScriptRoot }
    } Until (Test-Path -Path $Path)
    if ((Get-Item -Path $Path) -is [System.IO.DirectoryInfo]) {
        $SaveFileBrowser.InitialDirectory = Split-Path -Path $Path -Leaf
        $SaveFileBrowser.FileName=""
    } else {
        $SaveFileBrowser.InitialDirectory = Split-Path -Path $Path -Parent
        $SaveFileBrowser.FileName=Split-Path -Path $Path -Leaf
    }
    if ($SaveFileBrowser.ShowDialog() -ne "OK") {
        Write-Host "User Cancled" -ForegroundColor Red
        Write-Host
        Write-Host "[EXITING]" -ForegroundColor Red
        return
    }
    $IndexFile = $SaveFileBrowser.FileName
} else {
    if (-not (Test-Path $IndexFile -ErrorAction SilentlyContinue)) {
        Write-Host "Index File doesn't exist" -ForegroundColor Red
        Write-Host
        Write-Host "[EXITING]" -ForegroundColor Red
        return
    }
}   
$config.Config.IndexFile = $IndexFile
return;

#********************************************************
#Save the configuration file
#********************************************************
$ConfigFileContents=""
$ConfigFileContents+="@{`n"
$ConfigFileContents+="  DictionaryFile=`"$DictionaryFile`"`n"
$ConfigFileContents+="  IndexFile=`"$IndexFile`"`n"
$ConfigFileContents+="  SearchBase=`"$SearchBase`"`n"
$ConfigFileContents+="}"
write-host
Write-Host "NEW CONFIGURATION FILE" -BackgroundColor White -ForegroundColor Black
write-host "Writting Configuration" -ForegroundColor Green
Set-Content -Path $ConfigFile -Value  $ConfigFileContents
write-host
write-host "[COMPLETED]" -ForegroundColor Green