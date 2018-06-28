# =====================================================================================================
#                                          Licensing & Keys
# =====================================================================================================

$snagitLicense = 'ACFMR-BCMFH-YZAEW-KE78P-JC5C9'

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
#                                    App Installation Selection
# =====================================================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Optional App Selection Form'
$form.Size = New-Object System.Drawing.Size(350,600) #the outer box dimensions
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(130,525) #the "OK" button placement
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

<#
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,325) #the "Cancel" button placement
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)
#>

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Select all apps to install from the list below:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.Listbox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(315,20) #the inner box dimensions

$listBox.SelectionMode = 'MultiExtended'

[void] $listBox.Items.Add('( NONE )')
[void] $listBox.Items.Add('Adobe Acrobat DC')
[void] $listBox.Items.Add('Adobe Creative Cloud')
[void] $listBox.Items.Add('TO DO: Adobe Illustrator')
[void] $listBox.Items.Add('TO DO: Adobe InDesign')
[void] $listBox.Items.Add('TO DO: Adobe Photoshop')
[void] $listBox.Items.Add('TO DO: Adobe Photoshop Elements')
[void] $listBox.Items.Add('TO DO: Agresso Report Engine')
[void] $listBox.Items.Add('TO DO: Adobe Photoshop Elements')
[void] $listBox.Items.Add('TO DO: CD Generator')
[void] $listBox.Items.Add('TO DO: Carbonite')
[void] $listBox.Items.Add('Dropbox')
[void] $listBox.Items.Add('TO DO: ExpoCAD 7')
[void] $listBox.Items.Add('Filezilla')
[void] $listBox.Items.Add('TO DO: HTTP Watch')
[void] $listBox.Items.Add('iTunes')
[void] $listBox.Items.Add('Spotify')
[void] $listBox.Items.Add('TO DO: MathType 6.9 (Add-in)')
[void] $listBox.Items.Add('TO DO: Meeting Matrix 2010')
[void] $listBox.Items.Add('GoToMeeting')
[void] $listBox.Items.Add('Citrix Receiver')
[void] $listBox.Items.Add('TO DO: Microsoft Report Builder')
[void] $listBox.Items.Add('Notepad++')
[void] $listBox.Items.Add('TO DO: Outlook CRM 2015 Add-In')
[void] $listBox.Items.Add('R Statistical Analysis Software')
[void] $listBox.Items.Add('TO DO: SmartDraw')
[void] $listBox.Items.Add('VMware Workstation 14.1.2')
[void] $listBox.Items.Add('Dell Command Update')
[void] $listBox.Items.Add('Adobe Reader')
[void] $listBox.Items.Add('Lastpass for Browsers')
[void] $listBox.Items.Add('Lastpass for Applications')

$listBox.Height = 475 #the inner box height
$form.Controls.Add($listBox)
$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $apps = $listBox.SelectedItems
    $apps #prints out the list of appsw
}

# =====================================================================================================
#                                       Office Removal Tool
# =====================================================================================================

Write-Host "Run Office Removal Tool? [y/n] " -ForegroundColor Magenta -NoNewline
$input = Read-Host
    Switch ($input) 
     { 
       Y {Write-Host "Yes, Running Office Removal Tool...";  $RunOfficeRT=$true } 
       N {Write-Host "No, Skipping Office Removal Tool.";    $RunOfficeRT=$false} 
       Default {Write-Host "Running Office Removal Tool..."; $RunOfficeRT=$true }
     }


If ($RunOfficeRT -eq $true) {

    $OfficeRemovalToolDir = "\\prodfs01\Software-Public\PC_Software\Microsoft Office\Office Removal Tool"
    $OfficeRemovalTool = Start-Process -FilePath "$OfficeRemovalToolDir\o15-ctrremove.diagcab" -PassThru -Wait

    If (Test-Path $OfficeRemovalToolDir) {
        If ($OfficeRemovalTool.ExitCode -eq 0) {write-host "Office Removal Tool completed without errors"}
        else {"Office Removal Tool Failed"}
    }
    else {write-host "Installation for Office Removal Tool cannot be found"}
}

# =====================================================================================================
#                                     Required-SPIE-Repo App Install
# =====================================================================================================

Write-Host "Installing Webtrends..."
Start-Process 'C:\Windows\System32\msiexec.exe' -ArgumentList "/i \\prodfs01\DeployShare\WebtrendsReportExporter\ReportExporter.msi /norestart" -Wait	#removed the added /s
Write-Host "Webtrends Installed..."

Write-Host "Installing Office 365..."
Start-Process '\\prodfs01\DeployShare\Office365\Office365-64Bit.exe' -Verb runAs -Wait #removed the added -ArgumentList "/s"
Write-Host "Office 365 Installed..."

# =====================================================================================================
#                                       Chocolatety Install
# =====================================================================================================

Write-Host Installing Chocolatety...`n

Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
Start-Sleep -s 1
choco upgrade chocolatey
Start-Sleep -s 1
# Need to put repo on the NAS
#choco source add -n SPIErepo -s "\\mmcmurray-7050\c$\Users\michaelm\Documents\Github\choco\repo"

Write-Host Chocolatety Install Complete!`n

# =====================================================================================================
#                                 Required-Coco-Public-Repo App Install
# =====================================================================================================

Write-Host Installing Optionally Selected Apps...`n



# Public App List:
choco install dotnet3.5 -y
choco install flashplayerplugin -y
choco install google-chrome-for-enterprise -y
choco install firefoxesr -y
choco install jre8 --x86 -y
choco install snagit -params '/licenseCode:'$snagitLicense -y
choco install vlc -y
choco install 7zip.install -y



# =====================================================================================================
#                                         Optional App Install
# =====================================================================================================

if (('( NONE )' -Notin $apps) -Or (-Not ($apps.count -gt 0))) { #makes sure that the apps has something
    if ('Dell Command Update' -In $apps)             {choco install dellcommandupdate -y}
    if ('Adobe Reader' -In $apps)                    {choco install adobereader -y}
    if ('Notepad++' -In $apps)                       {choco choco install notepadplusplus -y}
    if ('Filezilla' -In $apps)                       {choco install filezilla -y}
    if ('iTunes' -In $apps)                          {choco install itunes -y}
    if ('Spotify' -In $apps)                         {choco install spotify -y}
    if ('Lastpass for Browsers' -In $apps)           {choco install lastpass -y}
    if ('Lastpass for Applications' -In $apps)       {choco install lastpass-for-applications -y}
    if ('Dropbox' -In $apps)                         {choco install dropbox -y}
    if ('Adobe Creative Cloud' -In $apps)            {choco install adobe-creative-cloud -y}
    if ('GoToMeeting' -In $apps)                     {choco install gotomeeting -y}
    if ('Citrix Receiver' -In $apps)                 {choco install citrix-receiver -y}
    if ('R Statistical Analysis Software' -In $apps) {choco install r.project -y}
    if ('VMware Workstation 14.1.2' -In $apps)       {choco install vmwareworkstation -y}
    if ('Adobe Acrobat DC' -In $apps)                {Start-Process '\\prodfs01\Software-Public\PC_Software\Adobe\Acrobat\Acrobat DC 2015\Acrobat_2015_Web_WWMUI.exe' -ArgumentList "/s" -Verb runAs -Wait}
    
}

Write-Host Optionally Selected App Install Complete!`n
