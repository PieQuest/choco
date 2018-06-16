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
@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

# =====================================================================================================
#                                       Coco-Public-Repo App Install
# =====================================================================================================

Write-Host Installing Coco-Public-Repo Apps...`n

#right now im just working on replicating the base app install list that we already have
#but only putting apps that we always need to keep up-to-date on there

# Public App List:
choco install flashplayerplugin -y
choco install googlechrome -y
choco install jre8 -y
choco install snagit /licenseCode:"ACFMR-BCMFH-YZAEW-KE78P-JC5C9"
choco install vlc -y
choco install 7zip -y
#choco install adobereader


Write-Host Coco-Public-Repo Install Complete!`n

# =====================================================================================================
#                                       Coco-Github-Repo App Install
# =====================================================================================================

Write-Host Installing Coco-Github-Repo Apps...`n

#choco install adobeacrobat  #make custom
choco install office365business #make this custom?