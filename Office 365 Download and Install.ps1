##########
# Office 365 Download and Installer
# Peter Grimley - https://github.com/PGrimley
##########

# Ask for elevated permissions if required
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}
 
##########
# Office 365
##########

##Download Office 2016 (O365) and store in Public Downloads Folder
$o365source = "http://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/O365BusinessRetail.img" 
$o365destination = "C:\Users\Public\Downloads\O365BusinessRetail.img" 
Invoke-WebRequest $o365source -OutFile $o365destination

#Mount O365 Image and Install O365 
$mountResult = Mount-DiskImage C:\Users\Public\Downloads\O365BusinessRetail.img
$mount = Mount-DiskImage C:\Users\Public\Downloads\O365BusinessRetail.img -PassThru
$driveLetter = ($mount | Get-Volume).DriveLetter

# Mount using New-PSDrive so other cmdlets in this session can see the new drive
New-PSDrive L -PSProvider FileSystem -Root "$($driveLetter):\"

cd L:
./setup.exe /install=agent /s 
