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
$Office2019Url      = $env:BUILD_Office2019Url
$Office2019XMLUrl   = $env:BUILD_Office2019XMLUrl

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


#=======================================================================
#   [OS] Install Office 365
#=======================================================================

<# ---------------------------------------------------------------------------
 Office installer – chooses Office 365 or Office 2019 based on BuildType.txt
    • Shared    → download + install Office 2019
    • Standard  → install Office 365
    • Rebuild   → install Office 365
    • Anything else → default to Office 365
---------------------------------------------------------------------------#>

# 1. URLs – adjust to your storage locations

# 2. Local working paths
$workingDir     = 'C:\Temp'
$localSetupPath = Join-Path $workingDir 'setup.exe'
$localXmlPath   = Join-Path $workingDir 'install.xml'

# 3. Ensure TLS 1.2 (needed in WinPE/WinRE for Azure blobs)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 4. Read build type, trimming CR/LF and possible UTF‑8 BOM
$buildTypeFile = 'C:\OSDCloud\BuildType.txt'
$buildType     = (Get-Content $buildTypeFile -Raw).Trim().Trim([char]0xFEFF)

# 5. Decide which package to use and whether we must pre‑download
$needsPreDownload = $false
switch -Regex ($buildType) {
    '^(?i)shared$' {
        $blobUrl          = $Office2019Url
        $xmlUrl           = $Office2019XMLUrl
        $message          = 'Install Office 2019 for Shared Machine'
        $needsPreDownload = $true           # Shared machines: stage source first
    }
    '^(?i)(standard|rebuild)$' {
        $blobUrl = $Office365Url
        $xmlUrl  = $Office365XMLUrl
        $message = 'Installing Office 365'
    }
    default {
        $blobUrl = $Office365Url
        $xmlUrl  = $Office365XMLUrl
        $message = "Can't determine build type, installing 365"
    }
}

Write-Host -ForegroundColor Cyan  "Build type detected: $buildType"
Write-Host -ForegroundColor Green $message

# 6. Ensure working directory exists
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
}

# 7. Download setup.exe and install.xml (‑UseBasicParsing avoids IE parser)
Invoke-WebRequest -Uri $blobUrl -OutFile $localSetupPath -UseBasicParsing
Invoke-WebRequest -Uri $xmlUrl  -OutFile $localXmlPath  -UseBasicParsing

# 8. Pre‑download source files only for Shared (Office 2019) machines
    #Start-Process -FilePath $localSetupPath `
                  #-ArgumentList "/download `"$downloadXmlPath`"" `
                  #-Wait -NoNewWindow
    # 8. Pre‑download source files only for Shared (Office 2019) machines
    Start-Process -FilePath $localSetupPath `
                  -ArgumentList "/configure `"$localXmlPath`"" `
                  -Wait -NoNewWindow




#=======================================================================
#   [OS] Install Company Portal
#=======================================================================

#Write-Host -ForegroundColor Green "Installing Company Portal"
#$packagePath    = "C:\OSDCloud\CompanyPortal\CompanyPortal.appxbundle"
#$dependencyPath = "C:\OSDCloud\CompanyPortal\Dependencies"
#$dependencies   = Get-ChildItem -Path $dependencyPath -Filter *.appx | ForEach-Object { $_.FullName }

#Add-AppxProvisionedPackage -Online -PackagePath $packagePath -DependencyPackagePath $dependencies -SkipLicense

#=======================================================================
#   [OS] Install Crowdstrike
#=======================================================================
          
Write-Host -ForegroundColor Green "Installing Crowdstrike Sensor from Azure"

# Variables
$blobUrl = $CrowdstrikeUrl
$localPath = "C:\Temp\FalconSensor_Windows.exe"

# Create download directory if it doesn't exist
if (-not (Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory
}

# Download setup.exe from Azure Blob
Invoke-WebRequest -Uri $blobUrl -OutFile $localPath


# Run the installer
Start-Process -FilePath $localPath -ArgumentList "/install /quiet /norestart /CID=$CrowdStrikeSecret" -Wait -NoNewWindow

       

#=======================================================================
#   [OS] Enroll in Autopilot
#=======================================================================

# Read Build Type and Builder

$BuildType = "C:\OSDCloud\BuildType.txt"

if (Test-Path $BuildType) {
    $GroupTagName = Get-Content "C:\OSDCloud\BuildType.txt" -Raw
    $GroupTag = $GroupTagName.Trim()
}
else {
    Write-Warning "Build Type file not found. Autopilot will register with Standard Group Tag ."
    $GroupTag = "Standard"
        
}

Write-Host -ForegroundColor Green "Starting Autopilot Registration"
        
import-module OSD 

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
# Define the full path to the AutoPilot script
$autoPilotScriptPath = "C:\OSDCloud\Scripts\Get-WindowsAutoPilotInfo.ps1"

# Prepare parameters for the Autopilot script
$AutopilotParams = @{
    Online               = $true
    TenantId             = $AutopilotTenantId
    AppId                = $AutopilotAppId
    AppSecret            = $AutopilotAppSecret
    GroupTag             = $GroupTag
    Assign               = $true
    AssignedComputerName = $deviceName
}

# Invoke the script file (this will load the function and execute it)
& $autoPilotScriptPath @AutopilotParams

# Display the equivalent command (if you still want to log it)
Write-Host -ForegroundColor Gray "Get-WindowsAutopilotInfo -Online -GroupTag $GroupTag -Assign -AssignedComputerName $deviceName"

        
        
#=======================================================================
#   [OS] Start Logic App
#=======================================================================

Write-Host -ForegroundColor Green "Starting Logic App for Whitelisting and Updating Jira"

# Send WebRequest to Logic App (IntuneDevices)

# Get the RAM and CPU Information and stick it in variables to pass to Logic App

# CPU Name
$cpuname = (Get-CimInstance Win32_Processor).Name

# Total RAM in GB
$ram = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)

# Disk Sizes in GB with " GB" added
$disksize = Get-CimInstance Win32_DiskDrive | ForEach-Object {
    "{0} GB" -f ([math]::Round($_.Size / 1GB, 2))
}

$model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model

$serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber


# Get the MAC address of the Wi-Fi adapter
$wifiMac = (Get-NetAdapter -Name WiFi | Select-Object -ExpandProperty MacAddress)

# Display the MAC address without the dash
$wifiMacWithoutDash = $wifiMac -replace '-', ''

Write-Host "Wi-Fi MAC Address: $wifiMacWithoutDash"


# Your Logic App HTTP trigger URL
$logicAppUrl = $LogicAppUrl

# Payload to send (can be empty if your Logic App doesn’t expect a body)
$payload = @{
    wifimac    = $wifiMacWithoutDash
    email      = "james.talbot@stmonicatrust.org.uk"
    deviceName = $deviceName
    cpuname = $cpuname
    ram = $ram
    disksize = $disksize
    model = $model
    serial = $serial
} | ConvertTo-Json -Depth 3

# Optional headers
$headers = @{
    "Content-Type" = "application/json"
}

# Invoke the Logic App
$response = Invoke-RestMethod -Uri $logicAppUrl -Method Post -Body $payload -Headers $headers

# Output response
$response


          
#=======================================================================
#   [OS] Tidy Up
#=======================================================================
Remove-Item -Path "C:\Temp" -Recurse -Force
