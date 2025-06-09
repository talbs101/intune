#=======================================================================
#   [OS] Obtain Secrets File
#=======================================================================

$secretsFile = 'C:\OSDCloud\Scripts\Secrets.ps1'
. $secretsFile
# ──────────────────────────────────────────────────────
# 2️⃣ Assign your “parameters” from the environment
# ──────────────────────────────────────────────────────
$SimpleHelpUrl      = $env:BUILD_SimpleHelpUrl
$Office365Url       = $env:BUILD_Office365Url
$Office365XMLUrl    = $env:BUILD_Office365XMLUrl
$CrowdStrikeUrl     = $env:BUILD_CrowdStrikeUrl
$CrowdStrikeSecret  = $env:BUILD_CrowdStrikeSecret
$LogicAppUrl        = $env:BUILD_LogicAppUrl
$AutopilotTenantId  = $env:BUILD_AutopilotTenantId
$AutopilotAppId     = $env:BUILD_AutopilotAppId
$AutopilotAppSecret = $env:BUILD_AutopilotAppSecret
#=======================================================================
#   [OS] Get Computer Name
#=======================================================================

# Read saved device name

$DeviceNameFile  = "C:\OSDCloud\DeviceName.txt"
$BuildTypeFile   = "C:\OSDCloud\BuildType.txt"
$BuilderFile     = "C:\OSDCloud\Builder.txt"



# ───────────────────────────────────────────────────────────────
# Rename the computer to the device name you collected
# ───────────────────────────────────────────────────────────────

$DeviceNameFile = "C:\OSDCloud\DeviceName.txt"

if (Test-Path $DeviceNameFile) {
    $deviceName = Get-Content "C:\OSDCloud\DeviceName.txt" -Raw
    $deviceName = $deviceName.Trim()
}
else {
    Write-Warning "DeviceName file not found. Autopilot will register with temporary name $env:computername ."
    
}


#=======================================================================
#   [OS] Decrypt BitLocker
#=======================================================================

# Decrypting no matter the current status so that Intune policy enforces Encryption to 256bit strength
Write-Host -ForegroundColor Green "Decrypting BitLocker"
Manage-bde -off C: 

#=======================================================================
#   [OS] Enable Location Services
#=======================================================================

# Required to fix the issue where Intune doesn't obtain the WiFi mac address
Write-Host -ForegroundColor Green "Enabling Location Services"

# Define the registry path
$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"

# Check if the registry key exists, if not, create it
if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}

# Set the registry value
Set-ItemProperty -Path $registryPath -Name "Value" -Type String -Value "Allow" 

#=======================================================================
#   [OS] Install SimpleHelp
#=======================================================================


Write-Host -ForegroundColor Green "Installing SimpleHelp from Azure"
# Check if SimpleHelp is already installed

       
# Define download URL
$blobUrl = $SimpleHelpUrl

# Define local download path
$downloadPath = "C:\Temp\Remote-Access-windows64-online.exe"

# Ensure C:\Temp exists
if (-Not (Test-Path -Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
}

# Download the file
Invoke-WebRequest -Uri $blobUrl -OutFile $downloadPath

# Define installation arguments as an array
$installArgs = @("/S", "/NAME=AUTODETECT", "/HOST=https://sh.stmonicatrust.org.uk:444", "/NOSHORTCUTS")

# Run the installer
Start-Process -FilePath $downloadPath -ArgumentList $installArgs -Wait -NoNewWindow


