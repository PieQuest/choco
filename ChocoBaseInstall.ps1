# =====================================================================================================
#                                      Set Execution Policy
# =====================================================================================================

# Testing this:

Set-ExecutionPolicy -ExecutionPolicy Unrestricted

$rstrct = New-Object -com wscript.shell;
#$rstrct.AppActivate('Windows PowerShell')
Start-Sleep -s 1
$rstrct.SendKeys('a{ENTER}')

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
#                                       Office Removal Tool
# =====================================================================================================

$OfficeRemovalToolDir = "\\prodfs01\Software-Public\PC_Software\Microsoft Office\Office Removal Tool"
$OfficeRemovalTool = Start-Process -FilePath "$OfficeRemovalToolDir\o15-ctrremove.diagcab" -PassThru -Wait

If (Test-Path $OfficeRemovalToolDir)
{
If ($OfficeRemovalTool.ExitCode -eq 0){
write-host "Office Removal Tool completed without errors"
} else {"Office Removal Tool Failed"}
}
else { write-host "Installation for Office Removal Tool cannot be found" }


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
choco source add -n SPIErepo -s \\mmcmurray-7050\c$\Users\michaelm\Documents\Github\choco\repo
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
choco install snagit /licenseCode:"ACFMR-BCMFH-YZAEW-KE78P-JC5C9" -y
choco install vlc -y
choco install 7zip -y
#choco install adobereader -y
#choco install notepadplusplus -y
#choco install quicktime -y
#choco install filezilla -y
#choco install itunes -y
#choco install spotify -y


Write-Host Coco-Public-Repo Install Complete!`n

# =====================================================================================================
#                                       Coco-SPIE-Repo App Install
# =====================================================================================================

Write-Host Installing Coco-SPIE-Repo Apps...`n

choco install office365 #change to SPIEoffice365
#choco install adobeacrobat  #make custom
#choco install office365business #make this custom?
#SPIE DropBox





# =====================================================================================================
#                                      Set Execution Policy
# =====================================================================================================

# Test:
Set-ExecutionPolicy -ExecutionPolicy Default
<#
$rstrct = New-Object -com wscript.shell;
#$rstrct.AppActivate('Windows PowerShell')
Start-Sleep -s 1
$rstrct.SendKeys('a{ENTER}')
#>
