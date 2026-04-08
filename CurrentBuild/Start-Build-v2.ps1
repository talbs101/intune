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
        [string]$Status    = "success",
        [string]$ErrorMsg  = "",
        [hashtable]$Extra  = @{}
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

        Write-Host "[$Stage] Event sent — $Status" -ForegroundColor Cyan

        # Return the response so the caller can use it
        return $response

    } catch {
        Write-Warning "[$Stage] Failed to send build event: $_"
        return $null
    }
}


#=======================================================================
#   [OS] Get Computer Name / Build Info
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

# Collect serial immediately — needed for ALL Send-BuildEvent calls
$serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
Write-Host "Serial: $serial" -ForegroundColor Gray

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
    # Define local download path
    $downloadPath = "C:\Temp\Remote-Access-windows64-online.exe"

    # Ensure C:\Temp exists
    if (-Not (Test-Path -Path "C:\Temp")) {
        New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
    }

    # Download the file
    Invoke-WebRequest -Uri $SimpleHelpUrl -OutFile $downloadPath

    # Run the installer
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
#   [OS] Stage: Create Jira Asset
#=======================================================================
#=======================================================================
#   [OS] Collect Hardware Info (used throughout)
#=======================================================================

$cpuName  = (Get-CimInstance Win32_Processor).Name
$ram      = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
$diskSize = Get-CimInstance Win32_DiskDrive | ForEach-Object { "{0} GB" -f ([math]::Round($_.Size / 1GB, 2)) }
$model    = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
# $serial already set above — don't re-declare here

$wifiAdapter = Get-NetAdapter | Where-Object {
    $_.Name -like "*Wi-Fi*" -or
    $_.InterfaceDescription -match "Wireless|Wi-Fi|802\.11|MediaTek|Intel.*WiFi|Qualcomm.*WiFi|Realtek.*WiFi"
} | Select-Object -First 1

if ($wifiAdapter) {
    $wifiMac = ($wifiAdapter.MacAddress) -replace '-', ''
    Write-Host "Wi-Fi MAC Address: $wifiMac" -ForegroundColor Green

    Send-BuildEvent -Stage "JiraAsset" -Status "success" -Extra @{
        cpuName  = $cpuName
        ram      = "$ram GB"
        diskSize = ($diskSize -join ", ")
        model    = $model
        wifiMac  = $wifiMac
    }

} else {
    $wifiMac = "NOT_FOUND"
    Write-Host "ERROR: No Wi-Fi adapter found!" -ForegroundColor Red

    Send-BuildEvent -Stage "JiraAsset" -Status "failed" -ErrorMsg "No Wi-Fi adapter found — MAC address could not be obtained. Jira asset may be incomplete."
}



#=======================================================================
#   [OS] Stage: BuildComplete
#   Final event — Logic App transitions Jira ticket, applies Meraki policy
#=======================================================================

Send-BuildEvent -Stage "BuildComplete"

#=======================================================================
#   [OS] Tidy Up
#=======================================================================

Remove-Item -Path "C:\Temp"      -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\OSDCloud\"  -Recurse -Force -ErrorAction SilentlyContinue
