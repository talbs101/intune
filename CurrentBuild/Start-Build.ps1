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

$DeviceNameFile = "C:\OSDCloud\DeviceName.txt"
$BuildTypeFile = "C:\OSDCloud\BuildType.txt"
$BuilderFile = "C:\OSDCloud\Builder.txt"

if (Test-Path $DeviceNameFile) {
    $deviceName = Get-Content "C:\OSDCloud\DeviceName.txt" -Raw
    $deviceName = $deviceName.Trim()
}
else {
    Write-Warning "DeviceName file not found. Autopilot will register with temporary name $env:computername ."
    
}

if (Test-Path $BuilderFile) {
    $builderName = Get-Content "C:\OSDCloud\Builder.txt" -Raw
    $builder= $buildName.Trim()
}
else {
    Write-Warning "Build Type file not found."
    $builder = "James"
    
}

# ───────────────────────────────────────────────────────────────
# Rename the computer to the device name you collected
# ───────────────────────────────────────────────────────────────

if (-not [string]::IsNullOrWhiteSpace($deviceName)) {
    Write-Host -ForegroundColor Green "Renaming computer to '$deviceName'..."
    try {
        Rename-Computer -NewName $deviceName -Force -PassThru | Write-Host
        Write-Host -ForegroundColor Green "Rename queued. Waiting for restart to apply new name..."
        
    }
    catch {
        Write-Warning "Failed to rename computer: $_"
    }
}
else {
    Write-Warning "No deviceName provided; skipping rename."
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
#   [OS] Install Windows Updates
#=======================================================================
    
Write-Host -ForegroundColor Green "Installing Windows Updates"
    
# How To: Update Windows using the PSWindowsUpdate Module

$UpdateWindows = $false
if (!(Get-Module PSWindowsUpdate -ListAvailable)) {
    try {
        Install-Module PSWindowsUpdate -Force
    }
    catch {
        Write-Warning 'Unable to install PSWindowsUpdate PowerShell Module'
        $UpdateWindows = $false
           
    }
}

if ($UpdateWindows) {
    Write-Host -ForegroundColor DarkCyan 'Add-WUServiceManager -MicrosoftUpdate -Confirm:$false'
    Add-WUServiceManager -MicrosoftUpdate -Confirm:$false

    Write-Host -ForegroundColor DarkCyan 'Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot'
    #Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -NotTitle 'Malicious'
}
   


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
          
Write-Host -ForegroundColor Green "Installing Office 365 from Azure"

# Variables
$blobUrl = $Office365Url
$localPath = "C:\Temp\setup.exe"
$installXmlPath = "C:\Temp\install.xml"

# Create download directory if it doesn't exist
if (-not (Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory
}

# Download setup.exe from Azure Blob
Invoke-WebRequest -Uri $blobUrl -OutFile $localPath

# Optionally download the install.xml if it's also in the blob
$xmlUrl = $Office365XMLUrl
Invoke-WebRequest -Uri $xmlUrl -OutFile $installXmlPath

# Run the installer
Start-Process -FilePath $localPath -ArgumentList "/configure `"$installXmlPath`"" -Wait -NoNewWindow

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
#Get-WindowsAutopilotInfo -Online -GroupTag Standard -Assign -AssignedComputerName $deviceName

$AutopilotParams = @{
    Online               = $true
    TenantId             = $AutopilotTenantId
    AppId                = $AutopilotAppId
    AppSecret            = $AutopilotAppSecret
    GroupTag             = $GroupTag
    Assign               = $true
    AssignedComputerName = $deviceName
}

Get-WindowsAutoPilotInfo @AutopilotParams

& $autoPilotScriptPath @AutopilotParams

write-host -ForegroundColor Gray '$AutopilotRegisterCommand'" = Get-WindowsAutopilotInfo -Online -GroupTag $GroupTag -Assign -AssignedComputerName $deviceName"
        
        
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
Read-Host -Prompt "Yes"

Remove-Item -Path "C:\Temp" -Recurse -Force
Write-Host -ForegroundColor DarkGray "Executing Cleanup Script"







