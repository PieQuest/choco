$ErrorActionPreference = 'Stop'; # stop on all errors
$toolsDir	= "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
$url	= '\\prodfs01\Software-Public\PC_Software\Microsoft Office\Office 365\Office365-64Bit.exe'
$fileLocation	= '\\prodfs01\Software-Public\PC_Software\Microsoft Office\Office 365\Office365-64Bit.exe'

$packageArgs = @{
  packageName   = 'office365'
  unzipLocation = $toolsDir
  fileType      = 'EXE' #only one of these: exe, msi, msu
  #url           = $url
  file         = $fileLocation

  softwareName  = 'office365' #part or all of the Display Name as you see it in Programs and Features. It should be enough to be unique

  # MSI
  silentArgs    = "/quiet /qn /norestart /l*v `"$($env:TEMP)\$($packageName).$($env:chocolateyPackageVersion).MsiInstall.log`"" # ALLUSERS=1 DISABLEDESKTOPSHORTCUT=1 ADDDESKTOPICON=0 ADDSTARTMENU=0
  validExitCodes= @(0, 3010, 1641)
}

Install-ChocolateyPackage @packageArgs