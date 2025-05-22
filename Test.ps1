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
$logicAppUrl = "https://prod-27.ukwest.logic.azure.com:443/workflows/622bc799ec1b4fd9974f92bea25b482d/triggers/When_a_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_a_HTTP_request_is_received%2Frun&sv=1.0&sig=2KYl03U2XnJyty913PD_M7bAwDafBHFOVfdIkMLOkvU"

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
#   [OS] Branding
#=======================================================================        

# ---------------- CONFIGURATION ----------------
# 1) Public URL of the blob
$blobUrl = 'https://intuneshareddata.blob.core.windows.net/intuneshared/Desktop3.jpg'

# 2) Where to stash the downloaded wallpaper
$destinationFolder = Join-Path $env:LOCALAPPDATA 'Temp\IntuneWallpapers'
$localFileName     = 'wallpaper.jpg'

# 3) Registry marker so we only run this once per user
$markerKeyPath     = 'HKCU:\Software\Contoso\WallpaperDeployment'
$markerValueName   = 'WallpaperSet'
# -----------------------------------------------

# Helper: write and exit
function Finish {
    param($message, $code)
    Write-Host $message
    exit $code
}

# If marker exists, bail out immediately
if ( Test-Path $markerKeyPath ) {
    $already = Get-ItemProperty -Path $markerKeyPath -Name $markerValueName -ErrorAction SilentlyContinue
    if ( $already.$markerValueName ) {
        Finish "✔ Wallpaper already deployed for this user; skipping." 0
    }
}

# Ensure folder exists
if ( -not (Test-Path $destinationFolder) ) {
    try {
        New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
    }
    catch {
        Finish "❌ Could not create folder $destinationFolder:`n  $_" 1
    }
}

# Download the image
$destinationPath = Join-Path $destinationFolder $localFileName
try {
    Invoke-WebRequest -Uri $blobUrl -OutFile $destinationPath -UseBasicParsing -ErrorAction Stop
}
catch {
    Finish "❌ Download failed:`n  $($_.Exception.Message)" 1
}

# P/Invoke to set the wallpaper
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
    [DllImport("user32.dll",SetLastError=true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@ -Language CSharp

# SPI_SETDESKWALLPAPER = 20; SPIF_UPDATEINIFILE | SPIF_SENDCHANGE = 3
$ok = [NativeMethods]::SystemParametersInfo(20, 0, $destinationPath, 3)
if ( -not $ok ) {
    $err = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    Finish "❌ Failed to set wallpaper (Win32 error $err)" 1
}

# Mark success in the registry
try {
    if ( -not (Test-Path $markerKeyPath) ) {
        New-Item -Path $markerKeyPath -Force | Out-Null
    }
    New-ItemProperty -Path $markerKeyPath `
                    -Name $markerValueName `
                    -Value (Get-Date -Format 's') `
                    -PropertyType String `
                    -Force | Out-Null
}
catch {
    Write-Warning "⚠️ Could not write registry marker: $($_.Exception.Message)"
    # We don't treat this as fatal—wallpaper is already set.
}

Finish "✅ Wallpaper applied and won't run again for this user." 0



          
#=======================================================================
#   [OS] Tidy Up
#=======================================================================

Remove-Item -Path "C:\Temp" -Recurse -Force



