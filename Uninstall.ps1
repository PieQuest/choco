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
#                                       Coco-Public-Repo App Uninstall
# =====================================================================================================

Write-Host Uninstalling Coco-Public-Repo Apps...`n

<#
right now im just working on replicating the base app install list that we already have
but only putting apps that we always need to keep up-to-date on there
#>

# Public App List:
choco uninstall flashplayerplugin -y
choco uninstall googlechrome -y
choco uninstall firefoxesr -y
choco uninstall jre8 -y
choco uninstall snagit -y
choco uninstall vlc -y
choco uninstall 7zip -y
#choco uninstall dellcommandupdate -y
#choco uninstall adobereader -y
#choco uninstall notepadplusplus -y
#choco uninstall quicktime -y
#choco uninstall filezilla -y
#choco uninstall itunes -y
#choco uninstall spotify -y
#choco uninstall lastpass -y
#choco uninstall lastpass-for-applications -y


Write-Host Coco-Public-Repo Uninstall Complete!`n

# =====================================================================================================
#                                       Chocolatety Uninstall
# =====================================================================================================

Write-Host Uninstalling Chocolatety...`n
choco uninstall chocolatey
Write-Host Chocolatety Uninstall Complete!`n

# =====================================================================================================
#                                           SPIE-Repo App Install
# =====================================================================================================

Write-Host Uninstalling SPIE-Repo Apps...`n

#Webtrends Report Exporter
$webtrends = Get-WmiObject Win32_Product -filter "Name='Webtrends Report Exporter 9.2'"
$webtrends.Uninstall()

#Adobe Acrobat DC
$acrobat = Get-WmiObject Win32_Product -filter "Name='Adobe Acrobat DC (2015)'"
$acrobat.Uninstall()

Write-Host SPIE-Repo Install Complete!`n


