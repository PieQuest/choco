# =====================================================================================================
#                                      Change Execution Policy
# =====================================================================================================

# Testing this:

#Set-ExecutionPolicy Unrestricted -Force -Scope Process
Set-ExecutionPolicy -ExecutionPolicy Unrestricted

$rstrct = New-Object -com wscript.shell;
#$rstrct.AppActivate('Windows PowerShell')
Start-Sleep -s 1
$rstrct.SendKeys('a{ENTER}')

# =====================================================================================================
#                  possible function to add users &/or groups to local admin group
# =====================================================================================================
<#
function Add-LocalUser{
     Param(
        $computer=$env:computername,
        $group='Guests',
        $userdomain=$env:userdomain,
        $username=$env:username
    )
        ([ADSI]"WinNT://$computer/$Group,group").psbase.Invoke("Add",([ADSI]"WinNT://$domain/$user").path)
}
#>
# =====================================================================================================
#                                   .NET Framework Redundancy Check
# =====================================================================================================

#Enabling .NET Framework in case someone forgets to enable it during initial setup. .Net Framework is required by some programs 
#installed by this script. This step is just to ensure no installers crash. It mounts the Windows 10 disk, runs a .bat to enable
#.NET Framework (.bat file provided by Amilia Mulka), then dismounts the disk image.

Write-Host 'Enabling .NET Framework 3.5...'`n
Mount-DiskImage '\\prodfs01\Software-Public\Operating Systems\Windows X Enterprise\ISO\SW_DVD5_WIN_ENT_10_1607_64BIT_English_MLF_X21-07102.ISO'
Start-Process '\\prodfs01\Software\Desktop Support Group\Scripts\netframeworkfix.bat' -Wait
Dismount-DiskImage '\\prodfs01\Software-Public\Operating Systems\Windows X Enterprise\ISO\SW_DVD5_WIN_ENT_10_1607_64BIT_English_MLF_X21-07102.ISO'
Write-Host '.NET Framework 3.5 Enabled.'`n

# =====================================================================================================
#                                       Office Removal Tool
# =====================================================================================================

Write-Host Running Office Removal Tool...`n

$OfficeRemovalToolDir = "\\prodfs01\Software-Public\PC_Software\Microsoft Office\Office Removal Tool"
$OfficeRemovalTool = Start-Process -FilePath "$OfficeRemovalToolDir\o15-ctrremove.diagcab" -PassThru -Wait

If (Test-Path $OfficeRemovalToolDir)
{
If ($OfficeRemovalTool.ExitCode -eq 0) {write-host "Office Removal Tool completed without errors"}
else {"Office Removal Tool Failed"}
}
else {write-host "Installation for Office Removal Tool cannot be found"}

# =====================================================================================================
#                                       Chocolatety Install
# =====================================================================================================

Write-Host Installing Chocolatety...`n

<#
If (Get-ExecutionPolicy -eq Restricted)
{
Set-ExecutionPolicy AllSigned
}
#>
Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
Start-Sleep -s 1
choco upgrade chocolatey
Start-Sleep -s 1
# Need to put repo on the NAS
#choco source add -n SPIErepo -s "\\mmcmurray-7050\c$\Users\michaelm\Documents\Github\choco\repo"

Write-Host Chocolatety Install Complete!`n

# =====================================================================================================
#                                       Coco-Public-Repo App Install
# =====================================================================================================

Write-Host Installing Coco-Public-Repo Apps...`n

<#
right now im just working on replicating the base app install list that we already have
but only putting apps that we always need to keep up-to-date on there
#>

# Public App List:
choco install flashplayerplugin -y
choco install googlechrome -y
choco install firefoxesr -y
choco install jre8 -y
choco install snagit -params '/licenseCode:ACFMR-BCMFH-YZAEW-KE78P-JC5C9'
choco install vlc -y
choco install 7zip -y
#choco install dellcommandupdate -y
#choco install adobereader -y
#choco install notepadplusplus -y
#choco install quicktime -y
#choco install filezilla -y
#choco install itunes -y
#choco install spotify -y
#choco install lastpass -y
#choco install lastpass-for-applications -y


Write-Host Coco-Public-Repo Install Complete!`n

# =====================================================================================================
#                                           SPIE-Repo App Install
# =====================================================================================================

Write-Host Installing SPIE-Repo Apps...`n

#Webtrends Report Exporter
Start-Process 'C:\Windows\System32\msiexec.exe' -ArgumentList "/i \\prodfs01\DeployShare\WebtrendsReportExporter\ReportExporter.msi /s /norestart" -Wait	#added /s

#Office 365
Start-Process '\\prodfs01\DeployShare\Office365\Office365-64Bit.exe' -ArgumentList "/s" -Verb runAs -Wait #added -ArgumentList "/s"

#Adobe Acrobat DC
#Start-Process '\\prodfs01\Software-Public\PC_Software\Adobe\Acrobat\Acrobat DC 2015\Acrobat_2015_Web_WWMUI.exe' -ArgumentList "/s" -Verb runAs -Wait

#choco install office365 #change to SPIEoffice365
#choco install adobeacrobat  #make custom
#choco install office365business #make this custom?
#SPIE DropBox

Write-Host SPIE-Repo Install Complete!`n

# =====================================================================================================
#                                      Reset Execution Policy
# =====================================================================================================

# Test:
Set-ExecutionPolicy -ExecutionPolicy Default
<#
$rstrct = New-Object -com wscript.shell;
#$rstrct.AppActivate('Windows PowerShell')
Start-Sleep -s 1
$rstrct.SendKeys('a{ENTER}')
#>
