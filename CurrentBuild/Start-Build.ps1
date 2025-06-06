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

#region Read DeviceName
if (Test-Path $DeviceNameFile) {
    $deviceName = (Get-Content $DeviceNameFile -Raw).Trim()
}
else {
    Write-Warning "DeviceName file not found. Autopilot will register with temporary name '$env:COMPUTERNAME'."
    $deviceName = $env:COMPUTERNAME
}
#endregion

#region Read BuildType
if (Test-Path $BuildTypeFile) {
    $buildType = (Get-Content $BuildTypeFile -Raw).Trim()
}
else {
    Write-Warning "BuildType file not found. Defaulting BuildType to 'Standard'."
    $buildType = "Standard"
}
#endregion

#region Read Builder
if (Test-Path $BuilderFile) {
    $builder = (Get-Content $BuilderFile -Raw).Trim()
}
else {
    Write-Warning "Builder file not found. Defaulting Builder to 'James'."
    $builder = "James"
}
#endregion

# (Optional) Output what we ended up with:
Write-Output "DeviceName = $deviceName"
Write-Output "BuildType  = $buildType"
Write-Output "Builder    = $builder"


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
#   [OS] Install Company Portal
#=======================================================================
<#
    Install-CompanyPortal-OneByOne.ps1

    • Assumes it is already running with full Administrator (or SYSTEM) rights.
    • Downloads each file listed below from:
         https://intuneshareddata.blob.core.windows.net/intuneshared/apps/company-portal/
      into C:\OSDCloud\CompanyPortal\[subfolders…]
    • Then installs CompanyPortal.appxbundle + all *.appx dependencies.

    USAGE (unattended/OOBE):
      1. Copy this .ps1 onto your WinPE/MDT/SCCM share or local image.
      2. In your Task Sequence (MDT/SCCM), add a step:
         "Run PowerShell Script" → point at this file, Execution Policy = Bypass.
      3. Make sure the TS step runs as SYSTEM (default) or an already-elevated Admin.
      4. No UAC‐elevation code remains in this file—fully silent/unattended.
#>

# ----------------------------------------
# 1) VARIABLES
# ----------------------------------------

# Base URL for every file
$baseUrl = "https://intuneshareddata.blob.core.windows.net/intuneshared/Apps/Company-Portal"

# Local folder where we want to drop everything
$destinationRoot = "C:\OSDCloud\CompanyPortal"

# A hard‐coded list of all blob‐paths (relative to the container root “apps/company-portal/”):
#   – Root-level:                                “CompanyPortal.appxbundle”
#   – Dependencies folder:                       “Dependencies/<filename>”
#
# NOTE: Adjust any filenames here if yours differ slightly.
$allFiles = @(
    # 1) The main bundle:
    "CompanyPortal.appxbundle",
    "install.ps1",

    # 2) Under Dependencies\:
    "Dependencies/AUMIDs.txt",
    "Dependencies/MPAP_c797dbb4414543f59d35e59e5225824e_001.provxml",
    "Dependencies/Microsoft.NET.Native.Framework.2.2_2.2.29512.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies/Microsoft.NET.Native.Runtime.2.2_2.2.28604.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies/Microsoft.Services.Store.Engagement_10.0.23012.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies/Microsoft.UI.Xaml.2.7_7.2409.9001.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies/Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies/c797dbb4414543f59d35e59e5225824e_License1.xml"
)

# ----------------------------------------
# 2) ENSURE DESTINATION FOLDER EXISTS
# ----------------------------------------
if (-not (Test-Path -Path $destinationRoot -PathType Container)) {
    Write-Host "Creating folder: $destinationRoot"
    New-Item -Path $destinationRoot -ItemType Directory -Force | Out-Null
}
# ----------------------------------------
# 3) DOWNLOAD EACH FILE ONE-BY-ONE
# ----------------------------------------
foreach ($relativePath in $allFiles) {
    # 3a) Build the full URL (no spaces → no %20 needed here)
    $fileUrl = "$baseUrl/$relativePath"

    # 3b) Build the local path where we’ll save it
    $localFullPath = Join-Path $destinationRoot $relativePath
    $localDir      = Split-Path $localFullPath -Parent

    # 3c) Create subfolder if it doesn't exist
    if (-not (Test-Path -Path $localDir -PathType Container)) {
        Write-Host " → Creating subfolder: $localDir"
        New-Item -Path $localDir -ItemType Directory -Force | Out-Null
    }

    # 3d) Download only if it’s not already present, or if you want to overwrite, you can remove the -ErrorAction check
    if (-not (Test-Path -Path $localFullPath -PathType Leaf)) {
        Write-Host " ↓ Downloading: $relativePath"
        try {
            Invoke-WebRequest -Uri $fileUrl -OutFile $localFullPath -UseBasicParsing -ErrorAction Stop
        }
        catch {
            Write-Error "   !! Failed to download '$relativePath' from '$fileUrl': $_"
            Exit 1
        }
    }
    else {
        Write-Host "   (Already exists) Skipping: $relativePath"
    }
}

Write-Host ""
Write-Host "✔ All requested files have been downloaded into '$destinationRoot'."
Write-Host ""

# ----------------------------------------
# 4) INSTALL THE COMPANY PORTAL APPX BUNDLE + DEPENDENCIES
# ----------------------------------------
Write-Host "Installing CompanyPortal.appxbundle + dependencies..."

# 4a) Path to main bundle
$bundlePath = Join-Path $destinationRoot "CompanyPortal.appxbundle"

# 4b) Path to Dependencies folder
$dependencyFolder = Join-Path $destinationRoot "Dependencies"

# 4c) Verify the bundle is present
if (-not (Test-Path -Path $bundlePath -PathType Leaf)) {
    Write-Error "CompanyPortal.appxbundle not found at '$bundlePath'. Cannot proceed."
    Exit 1
}

# 4d) Gather all *.appx in Dependencies\
$dependencyAppxFiles = @()
if (Test-Path -Path $dependencyFolder -PathType Container) {
    $dependencyAppxFiles = Get-ChildItem -Path $dependencyFolder -Filter *.appx -File |
                           ForEach-Object { $_.FullName }
}

# 4e) Run Add-AppxProvisionedPackage
try {
    Add-AppxProvisionedPackage -Online `
        -PackagePath $bundlePath `
        -DependencyPackagePath $dependencyAppxFiles `
        -SkipLicense

    Write-Host "✔ Company Portal installation succeeded."
}
catch {
    Write-Error "   !! Failed to install Company Portal: $_"
    #Exit 1
}

Write-Host ""
Write-Host "All done. Exiting."
#Exit 0

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
