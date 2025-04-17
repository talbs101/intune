Write-Output "--------------------------------------"
Write-Output ""
Write-Output "OSDCloud Apply OS Step"
Write-Output ""
#Set OSDCloud Params
$OSName = "Windows 11 24H2 x64"
Write-Output "OSName: $OSName"
$OSEdition = "Pro"
Write-Output "OSEdition: $OSEdition"
$OSActivation = "Volume"
Write-Output "OSActivation: $OSActivation"
$OSLanguage = "en-gb"
Write-Output "OSLanguage: $OSLanguage"





#Launch OSDCloud
Write-Output "Launching OSDCloud"
Write-Output ""
Write-Output "Start-OSDCloud -ZTI -OSName $OSName -OSEdition $OSEdition -OSActivation $OSActivation -OSLanguage $OSLanguage"
Write-Output ""
Start-OSDCloud -ZTI -OSName $OSName -OSEdition $OSEdition -OSActivation $OSActivation -OSLanguage $OSLanguage 
Write-Output ""
Write-Output "--------------------------------------"

#Copy new SetupComplete.cmd files - potentially host that in Github.
Copy-Item "X:\OSDCloud\Config\Scripts\SetupComplete\Build.ps1" "C:\Windows\Setup\scripts\Build.ps1" -Force
#Copy-Item "X:\SetupComplete.cmd" "C:\OSDCloud\Scripts\SetupComplete\SetupComplete.cmd" -Force

#Copy-Item "X:\Build.ps1" "C:\Windows\Setup\Scripts\Build.ps1" -Force
#Copy-Item "X:\SetupComplete.cmd" "C:\Windows\Setup\Scripts\SetupComplete.cmd" -Force
#Copy Custom Scripts to run in OOBE.

Write-Host -ForegroundColor Green "Staging post-install scripts..."

# Write oobe.cmd
$OOBECMD = @'
@echo off
start /wait powershell.exe -NoL -ExecutionPolicy Bypass -F C:\Windows\Setup\Scripts\Build.ps1
exit
'@

$OOBECMD | Out-File -FilePath 'C:\Windows\Setup\Scripts\oobe.cmd' -Encoding ascii -Force

# SetupComplete.cmd that launches oobe.cmd
@"
@echo off
start /wait C:\Windows\Setup\Scripts\oobe.cmd
exit
"@ | Out-File -FilePath 'C:\Windows\Setup\Scripts\SetupComplete.cmd' -Encoding ascii -Force



