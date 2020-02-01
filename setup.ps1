clear

#********************************************************
#Constants
#********************************************************
$ConfigFile="$PSScriptRoot\config.psd1"
$IndexPrecision=4 #This needs to match the precision useds in the 

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
        return;
    }
    Install-Module DSInternals
    If ((Get-InstalledModule -Name "DSInternals" -MinimumVersion 4.0 -ErrorAction SilentlyContinue))  {
        write-host "[Installed]" -ForegroundColor Green
    } else {
        Write-Host "Unable to Install DSInternals due to error" -ForegroundColor Red
        Write-Host "Please correct the issue and try again" -ForegroundColor Red
        Write-Host
        Write-Host "[EXITING]" -ForegroundColor Red
        return;
    }
}


#********************************************************
#Load the configuration file - If it exists
#           .\config.psd1
#It contains the location of the dictionary, index, and basic settings
#********************************************************
Write-Host
Write-Host "PREVIOUS CONFIGURATION" -BackgroundColor White -ForegroundColor Black
if (Test-Path $ConfigFile -ErrorAction SilentlyContinue) {
    Write-Host "Configuration Loaded" -ForegroundColor Green
    $Config=Import-PowerShellDataFile -Path $ConfigFile
} else {
    Write-Host "Configuration Doesn't exist" -ForegroundColor Red
    $Config=@{}
}
$SearchBase=[string]$Config.SearchBase
$DictionaryFile=[string]$Config.DictionaryFile
$IndexFile=[string]$Config.IndexFile


#********************************************************
#Get the domain and OU of where the user accounts are
#Example: OU=Users,DC=Domain,DC=com
#********************************************************
Write-Host
Write-Host "SEARCH BASE" -BackgroundColor White -ForegroundColor Black
Write-Host "Please enter the location of your users" -ForegroundColor Yellow
Write-Host "Example: OU=Users,DC=Domain,DC=com" -ForegroundColor Yellow
If ($SearchBase -ne "") {
    $NewSearchBase=Read-Host "Search Base [$SearchBase]"
} else {
    $NewSearchBase=Read-Host "Search Base"
}
if ($NewSearchBase -eq "") {$NewSearchBase=$SearchBase}
if ($NewSearchBase -eq "") {
    Write-Host "User Cancled" -ForegroundColor Red
    Write-Host
    Write-Host "[EXITING]" -ForegroundColor Red
    return
}
If ($NewSearchBase -like $SearchBase) {
    Write-Host "Search Base left unchanged" -ForegroundColor Green
} else {
    Write-Host "Search Base changed to: $NewSearchBase" -ForegroundColor Green
}
$SearchBase=$NewSearchBase


#********************************************************
#Get the domain and OU of where the user accounts are
#Example: OU=Users,DC=Domain,DC=com
#*******************************************************
Write-Host
Write-Host "BASIC OPTIONS" -BackgroundColor White -ForegroundColor Black


#********************************************************
#Get the location of the password dictionary
#Downloaded from: https://haveibeenpwned.com/Passwords
#********************************************************
Write-Host
Write-Host "DICTIONARY FILE" -BackgroundColor White -ForegroundColor Black
if ($DictionaryFile -ne "") {
    Write-Host "Password Dictionary File is currently set to: $DictionaryFile" -ForegroundColor Green
    if (-not (Test-Path $DictionaryFile -ErrorAction SilentlyContinue)) {
        Write-Host "File no longer exists" -ForegroundColor Red
        $DictionaryFile=""
    }
}
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    Title="Ordered NTLM Password Dictionary from https://haveibeenpwned.com/Passwords"
    InitialDirectory = $PSScriptRoot
    Filter = 'Password Dictionary (*.txt)|*.txt|All Files (*.*)|*.*'
}
if ($DictionaryFile -ne "") {
    $FileBrowser.InitialDirectory=Split-Path -Path $DictionaryFile -Parent
    $FileBrowser.FileName=Split-Path -Path $DictionaryFile -Leaf
}
Write-Host "Prompting for location of the Password Dictionary" -ForegroundColor Yellow
if ($FileBrowser.ShowDialog() -eq "OK") {
    if ($DictionaryFile -like $FileBrowser.FileName) {
        Write-Host "Dictionary File unchanged" -ForegroundColor Green
    } else {
        Write-Host "Dictionary File changed to: $($FileBrowser.FileName)" -ForegroundColor Green
    }
    $DictionaryFile = $FileBrowser.FileName
} else {
    Write-Host "User Cancled" -ForegroundColor Red
    Write-Host
    Write-Host "[EXITING]" -ForegroundColor Red
    return
}


#********************************************************
#Get the location of the index file
#The index file contains references to the location of hashes in the dictionary
#Required to speed up password searches
#********************************************************
Write-Host
Write-Host "INDEX FILE" -BackgroundColor White -ForegroundColor Black
if ($IndexFile -ne "") {
    Write-Host "Index File is currently set to: $IndexFile" -ForegroundColor Green
}
$FileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{ 
    Title="Index File"
    InitialDirectory = $PSScriptRoot
    Filter = 'Index File (*.index)|*.index|All Files (*.*)|*.*'
}
if ($IndexFile -ne "") {
    $FileBrowser.InitialDirectory=Split-Path -Path $IndexFile -Parent
    $FileBrowser.FileName=Split-Path -Path $IndexFile -Leaf
}
Write-Host "Prompting for location of the Index File" -ForegroundColor Yellow
if ($FileBrowser.ShowDialog() -eq "OK") {
    if ($IndexFile -like $FileBrowser.FileName) {
        Write-Host "Index File unchanged" -ForegroundColor Green
    } else {
        Write-Host "Index File changed to: $($FileBrowser.FileName)" -ForegroundColor Green
    }
    $IndexFile = $FileBrowser.FileName
} else {
    Write-Host "User Cancled" -ForegroundColor Red
    Write-Host
    Write-Host "[EXITING]" -ForegroundColor Red
    return
}


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