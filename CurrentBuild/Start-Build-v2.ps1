#=======================================================================
#   [OS] Obtain Secrets File
#=======================================================================

$secretsFile = 'C:\OSDCloud\Scripts\Secrets.ps1'
. $secretsFile

$SimpleHelpUrl      = $env:BUILD_SimpleHelpUrl
$Office365Url       = $env:BUILD_Office365Url
$Office365XMLUrl    = $env:BUILD_Office365XMLUrl
$CrowdStrikeUrl     = $env:BUILD_CrowdStrikeUrl
$CrowdStrikeSecret  = $env:BUILD_CrowdStrikeSecret
$LogicAppUrl        = $env:BUILD_LogicAppUrl2
$AutopilotTenantId  = $env:BUILD_AutopilotTenantId
$AutopilotAppId     = $env:BUILD_AutopilotAppId
$AutopilotAppSecret = $env:BUILD_AutopilotAppSecret
$Office2019Url      = $env:BUILD_Office2019Url
$Office2019XMLUrl   = $env:BUILD_Office2019XMLUrl

#=======================================================================
#   [OS] Send-BuildEvent Helper Function
#=======================================================================

function Send-BuildEvent {
    param(
        [string]$Stage,
        [string]$Status   = "success",
        [string]$ErrorMsg = "",
        [hashtable]$Extra = @{}
    )

    $base = @{
        stage     = $Stage
        status    = $Status
        error     = $ErrorMsg
        hostname  = $deviceName
        serial    = $serial
        buildType = $buildType
        builder   = $builder
        timestamp = (Get-Date -Format "o")
    }

    foreach ($key in $Extra.Keys) { $base[$key] = $Extra[$key] }

    $payload = $base | ConvertTo-Json -Depth 3

    try {
        $response = Invoke-RestMethod -Uri $LogicAppUrl -Method Post -Body $payload `
            -ContentType "application/json" -ErrorAction Stop

        Write-Host "[$Stage] Event sent - $Status" -ForegroundColor Cyan
        return $response

    } catch {
        Write-Warning "[$Stage] Failed to send build event: $_"
        return $null
    }
}

#=======================================================================
#   [OS] Get Computer Name / Build Info
#   Must come before hardware collection so $deviceName, $serial,
#   $buildType and $builder are all set before Send-BuildEvent is called
#=======================================================================

$DeviceNameFile = "C:\OSDCloud\DeviceName.txt"
$BuildTypeFile  = "C:\OSDCloud\BuildType.txt"
$BuilderFile    = "C:\OSDCloud\Builder.txt"

if (Test-Path $DeviceNameFile) {
    $deviceName = (Get-Content $DeviceNameFile -Raw).Trim()
} else {
    Write-Warning "DeviceName file not found. Using $env:COMPUTERNAME"
    $deviceName = $env:COMPUTERNAME
}

$buildType = if (Test-Path $BuildTypeFile) { (Get-Content $BuildTypeFile -Raw).Trim().Trim([char]0xFEFF) } else { "Standard" }
$builder   = if (Test-Path $BuilderFile)   { (Get-Content $BuilderFile   -Raw).Trim().Trim([char]0xFEFF) } else { "Unknown" }

$serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
Write-Host "Serial    : $serial"    -ForegroundColor Gray
Write-Host "Device    : $deviceName" -ForegroundColor Gray
Write-Host "Build Type: $buildType"  -ForegroundColor Gray
Write-Host "Builder   : $builder"    -ForegroundColor Gray

#=======================================================================
#   [OS] Collect Hardware Info
#   Collected once here - available to all stages below including
#   Meraki (needs wifiMac) and JiraAsset (needs all fields)
#=======================================================================

$cpuName  = (Get-CimInstance Win32_Processor).Name
$ram      = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
$diskSize = (Get-CimInstance Win32_DiskDrive | ForEach-Object { "{0} GB" -f ([math]::Round($_.Size / 1GB, 2)) }) -join ", "
$model    = (Get-CimInstance -ClassName Win32_ComputerSystem).Model

$wifiAdapter = Get-NetAdapter | Where-Object {
    $_.Name -like "*Wi-Fi*" -or
    $_.InterfaceDescription -match "Wireless|Wi-Fi|802\.11|MediaTek|Intel.*WiFi|Qualcomm.*WiFi|Realtek.*WiFi"
} | Select-Object -First 1

if ($wifiAdapter -and -not [string]::IsNullOrWhiteSpace($wifiAdapter.MacAddress)) {
    $wifiMac = ($wifiAdapter.MacAddress) -replace '-', ''
    Write-Host "Wi-Fi MAC : $wifiMac" -ForegroundColor Green
} else {
    $wifiMac = "NOT_FOUND"
    Write-Host "ERROR: No Wi-Fi adapter found or MAC address is empty!" -ForegroundColor Red
    Write-Host "Adapter found: $($wifiAdapter -ne $null)" -ForegroundColor Yellow
    if ($wifiAdapter) {
        Write-Host "Adapter name: $($wifiAdapter.Name)" -ForegroundColor Yellow
        Write-Host "Adapter MAC : '$($wifiAdapter.MacAddress)'" -ForegroundColor Yellow
    }
}

Write-Host "CPU       : $cpuName"  -ForegroundColor Gray
Write-Host "RAM       : $ram GB"   -ForegroundColor Gray
Write-Host "Disk      : $diskSize" -ForegroundColor Gray
Write-Host "Model     : $model"    -ForegroundColor Gray
Write-Host "Wi-Fi MAC : $wifiMac"  -ForegroundColor Gray

#=======================================================================
#   [OS] Decrypt BitLocker
#=======================================================================

Write-Host -ForegroundColor Green "Decrypting BitLocker"
Manage-bde -off C:

#=======================================================================
#   [OS] Enable Location Services
#=======================================================================

Write-Host -ForegroundColor Green "Enabling Location Services"
$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
if (-not (Test-Path $registryPath)) { New-Item -Path $registryPath -Force | Out-Null }
Set-ItemProperty -Path $registryPath -Name "Value" -Type String -Value "Allow"

#=======================================================================
#   [OS] Install SimpleHelp
#=======================================================================

Write-Host -ForegroundColor Green "Installing SimpleHelp from Azure"

try {
    $downloadPath = "C:\Temp\Remote-Access-windows64-online.exe"

    if (-Not (Test-Path -Path "C:\Temp")) {
        New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
    }

    Invoke-WebRequest -Uri $SimpleHelpUrl -OutFile $downloadPath

    Start-Process -FilePath $downloadPath -ArgumentList @(
        "/S", "/NAME=AUTODETECT",
        "/HOST=https://sh.stmonicatrust.org.uk:444",
        "/NOSHORTCUTS"
    ) -Wait -NoNewWindow

    Send-BuildEvent -Stage "SimpleHelpInstalled"

} catch {
    Send-BuildEvent -Stage "SimpleHelpInstalled" -Status "failed" -ErrorMsg $_.Exception.Message
    Write-Warning "SimpleHelp install failed: $_"
}

#=======================================================================
#   [OS] Install Office
#=======================================================================

Write-Host -ForegroundColor Cyan "Build type detected: $buildType"

try {
    $workingDir     = 'C:\Temp'
    $localSetupPath = Join-Path $workingDir 'setup.exe'
    $localXmlPath   = Join-Path $workingDir 'install.xml'

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (-not (Test-Path $workingDir)) { New-Item -Path $workingDir -ItemType Directory | Out-Null }

    switch -Regex ($buildType) {
        '^(?i)(care|kiosk)$' {
            $blobUrl = $Office2019Url
            $xmlUrl  = $Office2019XMLUrl
            Write-Host -ForegroundColor Green "Installing Office 2019 for Shared Machine"
        }
        default {
            $blobUrl = $Office365Url
            $xmlUrl  = $Office365XMLUrl
            Write-Host -ForegroundColor Green "Installing Office 365"
        }
    }

    Invoke-WebRequest -Uri $blobUrl -OutFile $localSetupPath -UseBasicParsing
    Invoke-WebRequest -Uri $xmlUrl  -OutFile $localXmlPath   -UseBasicParsing

    Start-Process -FilePath $localSetupPath -ArgumentList "/download `"$localXmlPath`"" -Wait -NoNewWindow
    Start-Process -FilePath $localSetupPath -ArgumentList "/configure `"$localXmlPath`"" -Wait -NoNewWindow

    Send-BuildEvent -Stage "OfficeInstalled" -Extra @{ officeType = $buildType }

} catch {
    Send-BuildEvent -Stage "OfficeInstalled" -Status "failed" -ErrorMsg $_.Exception.Message
    Write-Warning "Office install failed: $_"
}

#=======================================================================
#   [OS] Install CrowdStrike
#=======================================================================

Write-Host -ForegroundColor Green "Installing CrowdStrike Sensor from Azure"

try {
    $localPath = "C:\Temp\FalconSensor_Windows.exe"
    if (-not (Test-Path "C:\Temp")) { New-Item -Path "C:\Temp" -ItemType Directory }

    Invoke-WebRequest -Uri $CrowdStrikeUrl -OutFile $localPath
    Start-Process -FilePath $localPath `
        -ArgumentList "/install /quiet /norestart /CID=$CrowdStrikeSecret" `
        -Wait -NoNewWindow

    Send-BuildEvent -Stage "CrowdStrikeInstalled"

} catch {
    Send-BuildEvent -Stage "CrowdStrikeInstalled" -Status "failed" -ErrorMsg $_.Exception.Message
    Write-Warning "CrowdStrike install failed: $_"
}

#=======================================================================
#   [OS] Enroll in Autopilot
#=======================================================================

Write-Host -ForegroundColor Green "Starting Autopilot Registration"

try {
    $GroupTag = if (Test-Path $BuildTypeFile) {
        (Get-Content $BuildTypeFile -Raw).Trim()
    } else { "Standard" }

    Import-Module OSD
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $autoPilotScriptPath = "C:\OSDCloud\Scripts\Get-WindowsAutoPilotInfo.ps1"

    & $autoPilotScriptPath @{
        Online               = $true
        TenantId             = $AutopilotTenantId
        AppId                = $AutopilotAppId
        AppSecret            = $AutopilotAppSecret
        GroupTag             = $GroupTag
        Assign               = $true
        AssignedComputerName = $deviceName
    }

    Send-BuildEvent -Stage "AutopilotEnrolled" -Extra @{ groupTag = $GroupTag }

} catch {
    Send-BuildEvent -Stage "AutopilotEnrolled" -Status "failed" -ErrorMsg $_.Exception.Message
    Write-Warning "Autopilot enrolment failed: $_"
}

#=======================================================================
#   [OS] Stage: Meraki Whitelist
#   Sends MAC address to Logic App which applies the group policy
#=======================================================================

Write-Host -ForegroundColor Green "Sending Meraki whitelist event"

if ($wifiMac -ne "NOT_FOUND") {
    Send-BuildEvent -Stage "Meraki" -Extra @{ wifiMac = $wifiMac }
} else {
    Send-BuildEvent -Stage "Meraki" -Status "failed" -ErrorMsg "No Wi-Fi adapter found - cannot whitelist device in Meraki"
    Write-Warning "Meraki whitelist skipped - no Wi-Fi MAC address available"
}

#=======================================================================
#   [OS] Stage: Jira Asset
#   Sends full hardware payload to Logic App which creates the asset
#=======================================================================

Write-Host -ForegroundColor Green "Sending Jira asset creation event"

if ($wifiMac -ne "NOT_FOUND") {
    Send-BuildEvent -Stage "JiraAsset" -Status "success" -Extra @{
        cpuName  = $cpuName
        ram      = "$ram GB"
        diskSize = $diskSize
        model    = $model
        wifiMac  = $wifiMac
    }
} else {
    Send-BuildEvent -Stage "JiraAsset" -Status "failed" -ErrorMsg "No Wi-Fi adapter found - MAC address could not be obtained. Jira asset may be incomplete."
}

#=======================================================================
#   [OS] Stage: BuildComplete
#   Final event - Logic App transitions Jira ticket and sends email
#=======================================================================

Send-BuildEvent -Stage "BuildComplete"

#=======================================================================
#   [OS] Tidy Up
#=======================================================================

Remove-Item -Path "C:\Temp"        -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\OSDCloud\"   -Recurse -Force -ErrorAction SilentlyContinue
