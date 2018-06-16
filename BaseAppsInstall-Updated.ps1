# =========================================================================================================================================================
# Base Application Installation script
# Author: Alexander Covington
# 2/9/2017
#
# This script installs all of the base applications for standard SPIE desktops, excluding Sophos, Adobe Reader, OLE DB Provider, and FilesAnywhere.
# Installation files run by this script are either .msi or .exe installers. All of the .msi files are installed the same way:
#
#      Start-Process C:\Windows\System32\msiexec.exe -ArgumentList "/i \\prodfs01\DeployShare\Path\To\Install /qn /norestart" -Wait
#
# This script is intended to run from either the PowerShell shell or through the PowerShell ISE as an administrator. This script does not consistently 
# run when double-clicked. The script can be run by typing in the following command:
#
#      '\\prodfs01\Software\Desktop Support Group\Scripts\Base App Installer\BaseAppsInstall.ps1'
#
# Execution policy may need to be modified before running the script.
#
# Using msiexec, Windows is able to install all .msi files silently without any user interaction. It is given an argument list containing /i to indicate
# an install, path to the installer (all installers are located in the DeployShare folder), /qn to indicate a silent install with no user interaction, and
# /norestart to keep the computer from automatically restarting. The -Wait is used to run only one installer at a time.
#
# The installation process varies between the .exe files, however they all use a .bat file to run the correct installation sequence. All of the necessary
# .bat files are located on \\prodfs01\software\desktop support group\scripts\base app installer\. They also use the -Wait tag to only run one installation
# at a time. 
# ===========================================================================================================================================================

# =====================================================================================================
#                                   .NET Framework Redundancy Check
# =====================================================================================================

#Enabling .NET Framework in case someone forgets to enable it during initial setup. .Net Framework is required by some programs 
#installed by this script. This step is just to ensure no installers crash. It mounts the Windows 10 disk, runs a .bat to enable
#.NET Framework (.bat file provided by Amilia Mulka), then dismounts the disk image.
Write-Host Enabling .NET Framework...`n
Mount-DiskImage '\\prodfs01\Software-Public\Operating Systems\Windows X Enterprise\ISO\SW_DVD5_WIN_ENT_10_1607_64BIT_English_MLF_X21-07102.ISO'
Start-Process '\\prodfs01\Software\Desktop Support Group\Scripts\netframeworkfix.bat' -Wait
Dismount-DiskImage '\\prodfs01\Software-Public\Operating Systems\Windows X Enterprise\ISO\SW_DVD5_WIN_ENT_10_1607_64BIT_English_MLF_X21-07102.ISO'


# =====================================================================================================
#                                       Application Installation
# =====================================================================================================

#The Flash Player install is the only unique .msi install. Because Flash Player is updated so often, it scans the Flash Player
#directory and finds the folder that comes last alphabetically. It then finds the .msi file in that directory and uses that
#to install. 
Write-Host Installing Adobe Flash Player...`n
$flashItems = Get-ChildItem \\prodfs01\DeployShare\Adobe\Flash\
$adobeInstaller = Get-ChildItem $flashItems[($flashItems.length - 1)].FullName | where {$_.Name -Match "plugin"}
$adobeInstaller = $adobeInstaller.FullName 
Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList "/i $adobeInstaller /qn /norestart" -Wait

#Installing Chrome using standard .msi install method.
Write-Host Installing Chrome...`n
Start-Process C:\Windows\System32\msiexec.exe -ArgumentList "/i \\prodfs01\DeployShare\Chrome\Enterprise\googlechromestandaloneenterprise.msi /norestart" -Wait

#Installing Firefox by adding -ms as a command line argument
#It then copies the necessary files over to install the Authenticator plugin
Write-Host Installing Firefox...`n
Start-Process '\\prodfs01\DeployShare\Mozilla Firefox\Firefox52.7.3\setup.exe' -ArgumentList "-ms" -Wait
try{
    Copy-Item -path '\\prodfs01\DeployShare\Mozilla Firefox\Firefox Addons\distribution' -Recurse -Destination 'C:\Program Files (x86)\Mozilla Firefox'
    Copy-Item -Path '\\prodfs01\DeployShare\Mozilla Firefox\Firefox Addons\integrated_authentication_for_firefox-3.0.1-fx.xpi' -Destination 'C:\Program Files (x86)\Mozilla Firefox'
}
catch{
    Write-Host Unable to install add-in. Will have to be done manually
}

#Installing Java by running the Offline installer and passing /s through as an argument to indicate a silent install.
#Finding the Java installer by getting all of the items in the directory \\prodfs01\deployshare\sun java\current version. 
#There should only ever be one file in that directory, which is the Java installer.
Write-Host Installing Java...`n
$javaItems = Get-ChildItem '\\prodfs01\DeployShare\Sun Java\Current version'
Start-Process $javaItems.FullName -ArgumentList "/s" -Wait


#Installing Webtrends using standard .msi install method.
Write-Host Installing Webtrends Report Exporter...`n
Start-Process C:\Windows\System32\msiexec.exe -ArgumentList "/i \\prodfs01\DeployShare\WebtrendsReportExporter\ReportExporter.msi /norestart" -Wait

#Installing SnagIt using a .bat file. This is the only .bat file that needs to be told to run as an admin by using
#
#-Verb runAs
#
#The .bat file handles the activation of SnagIt.

Write-Host Installing SnagIt...`n
Start-Process \\prodfs01\DeployShare\SnagIT\Script\snagit.bat -Verb runAs -Wait


#Installing VLC using a .bat file.
Write-Host Installing VLC Media Player...`n
Start-Process '\\prodfs01\Software\Desktop Support Group\Scripts\Base App Installer\vlc.bat' -Verb runAs -Wait

#Installing 7-Zip
Write-Host Installing 7Zip...`n
Start-Process '\\prodfs01\DeployShare\7-Zip\7z-latest.exe' -Verb runAs -Wait

#Installing Microsoft Office 365
Write-Host Installing Office 365...`n
Start-Process '\\prodfs01\DeployShare\Office365\Office365-64Bit.exe' -Verb runAs -Wait

Write-Host Install Complete!
Pause