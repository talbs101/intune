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

#Copy new SetupComplete.cmd files
#Copy-Item "X:\Build.ps1" "C:\OSDCloud\Scripts\SetupComplete\Build.ps1" -Force
#Copy-Item "X:\SetupComplete.cmd" "C:\OSDCloud\Scripts\SetupComplete\SetupComplete.cmd" -Force

#Copy-Item "X:\Build.ps1" "C:\Windows\Setup\Scripts\Build.ps1" -Force
#Copy-Item "X:\SetupComplete.cmd" "C:\Windows\Setup\Scripts\SetupComplete.cmd" -Force
#Copy Custom Scripts to run in OOBE.


