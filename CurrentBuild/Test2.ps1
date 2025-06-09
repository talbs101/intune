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
#   [OS] Install Company Portal
#=======================================================================

# ----------------------------------------
# 1) VARIABLES
# ----------------------------------------

# Base URL for every file
$baseUrl = "https://intuneshareddata.blob.core.windows.net/intuneshared/Apps/Company-Portal"

# Local folder where we want to drop everything
$destinationRoot = "C:\OSDCloud\CompanyPortal"

# A hard‐coded list of all blob‐paths (relative to the container root "apps/company-portal/"):
#   – Root-level:                                "CompanyPortal.appxbundle"
#   – Dependencies folder:                       "Dependencies/<filename>"
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

    # 3d) Download only if it’s not already present
    if (-not (Test-Path -Path $localFullPath -PathType Leaf)) {
        Write-Host " ↓ Downloading: $relativePath"
        try {
            Invoke-WebRequest -Uri $fileUrl -OutFile $localFullPath -UseBasicParsing -ErrorAction Stop
        }
        catch {
            Write-Error "   !! Failed to download '$relativePath' from '$fileUrl': $_"
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
}

# 4d) Gather all *.appx in Dependencies\
$dependencyAppxFiles = @()
if (Test-Path -Path $dependencyFolder -PathType Container) {
    $dependencyAppxFiles = Get-ChildItem -Path $dependencyFolder -Filter *.appx -File |
                           ForEach-Object { $_.FullName }
}

# 4e) Run Add-AppxProvisionedPackage with splatting (no backticks)
$installParams = @{
    Online                  = $true
    PackagePath             = $bundlePath
    DependencyPackagePath   = $dependencyAppxFiles
    SkipLicense             = $true
}

try {
    Add-AppxProvisionedPackage @installParams
    Write-Host "✔ Company Portal installation succeeded."
}
catch {
    Write-Error "Failed to install Company Portal: $_"
}





