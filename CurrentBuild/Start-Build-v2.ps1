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
#   [OS] Send-LogEvent Helper Function
#   Sends error and warning details to Logic App for central logging.
#   Routes to BuildLog switch case which writes to BuildLogs table.
#=======================================================================

function Send-LogEvent {
    param(
        [string]$Level   = "ERROR",
        [string]$Section = "",
        [string]$Message = "",
        [string]$Detail  = ""
    )

    $payload = @{
        stage     = "BuildLog"
        level     = $Level
        section   = $Section
        message   = $Message
        detail    = $Detail
        hostname  = $deviceName
        serial    = $serial
        buildType = $buildType
        builder   = $builder
        timestamp = (Get-Date -Format "o")
    } | ConvertTo-Json -Depth 3

    try {
        Invoke-RestMethod -Uri $LogicAppUrl -Method Post -Body $payload `
            -ContentType "application/json" -ErrorAction Stop
        Write-Host "[BuildLog] $Level - $Section - $Message" -ForegroundColor DarkGray
    } catch {
        Write-Warning "[BuildLog] Failed to send log event: $_"
    }
}

#=======================================================================
#   [OS] Get Computer Name / Build Info
#   Explicit if/else blocks used instead of inline assignment
#   to avoid IEX misparse returning the file path as the value
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

if (Test-Path $BuildTypeFile) {
    $buildType = (Get-Content $BuildTypeFile -Raw).Trim().Trim([char]0xFEFF)
} else {
    $buildType = "Standard"
}

if (Test-Path $BuilderFile) {
    $builder = (Get-Content $BuilderFile -Raw).Trim().Trim([char]0xFEFF)
} else {
    $builder = "Unknown"
}

$serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber

Write-Host "Serial    : $serial"     -ForegroundColor Gray
Write-Host "Device    : $deviceName" -ForegroundColor Gray
Write-Host "Build Type: $buildType"  -ForegroundColor Gray
Write-Host "Builder   : $builder"    -ForegroundColor Gray

#=======================================================================
#   [OS] Decrypt BitLocker
#=======================================================================

Write-Host -ForegroundColor Green "Decrypting BitLocker"
Manage-bde -off C:

#=======================================================================
#   [OS] Enable Location Services
#   Required for Intune to obtain the WiFi MAC address
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
    $errMsg = $_.Exception.Message
    Send-BuildEvent -Stage "SimpleHelpInstalled" -Status "failed" -ErrorMsg $errMsg
    Send-LogEvent -Level "ERROR" -Section "SimpleHelpInstalled" `
        -Message "SimpleHelp install failed" -Detail $errMsg
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
    $errMsg = $_.Exception.Message
    Send-BuildEvent -Stage "OfficeInstalled" -Status "failed" -ErrorMsg $errMsg
    Send-LogEvent -Level "ERROR" -Section "OfficeInstalled" `
        -Message "Office install failed - BuildType: $buildType" -Detail $errMsg
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

    $cs = Start-Process -FilePath $localPath `
        -ArgumentList "/install /quiet /norestart /CID=$CrowdStrikeSecret" `
        -Wait -NoNewWindow -PassThru

    Write-Host "CrowdStrike exit code: $($cs.ExitCode)" -ForegroundColor Gray

    # Log non-zero exit codes as warnings even if we continue
    if ($cs.ExitCode -ne 0 -and $cs.ExitCode -ne 106) {
        Send-LogEvent -Level "WARNING" -Section "CrowdStrikeInstalled" `
            -Message "CrowdStrike returned unexpected exit code" `
            -Detail "Exit code: $($cs.ExitCode)"
    }

    # Wait for sensor to connect to Falcon cloud
    Write-Host "Waiting 60s for CrowdStrike cloud registration..." -ForegroundColor Gray
    Start-Sleep -Seconds 60

    # Verify service is running
    $csService = Get-Service -Name CSFalconService -ErrorAction SilentlyContinue
    if ($csService) {
        Write-Host "CSFalconService status: $($csService.Status)" -ForegroundColor Gray
        if ($csService.Status -ne "Running") {
            Send-LogEvent -Level "WARNING" -Section "CrowdStrikeInstalled" `
                -Message "CSFalconService not in Running state" `
                -Detail "Service status: $($csService.Status)"
        }
    } else {
        Send-LogEvent -Level "WARNING" -Section "CrowdStrikeInstalled" `
            -Message "CSFalconService not found after install" `
            -Detail "Service may not have registered correctly"
        Write-Host "WARNING: CSFalconService not found" -ForegroundColor Yellow
    }

    Send-BuildEvent -Stage "CrowdStrikeInstalled"

} catch {
    $errMsg = $_.Exception.Message
    Send-BuildEvent -Stage "CrowdStrikeInstalled" -Status "failed" -ErrorMsg $errMsg
    Send-LogEvent -Level "ERROR" -Section "CrowdStrikeInstalled" `
        -Message "CrowdStrike install failed" -Detail $errMsg
    Write-Warning "CrowdStrike install failed: $_"
}

#=======================================================================
#   [OS] Enroll in Autopilot
#=======================================================================

Write-Host -ForegroundColor Green "Starting Autopilot Registration"

try {
    $GroupTag = ""
    if (Test-Path $BuildTypeFile) {
        $GroupTag = (Get-Content $BuildTypeFile -Raw).Trim()
    } else {
        $GroupTag = "Standard"
    }

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
    $errMsg = $_.Exception.Message
    Send-BuildEvent -Stage "AutopilotEnrolled" -Status "failed" -ErrorMsg $errMsg
    Send-LogEvent -Level "ERROR" -Section "AutopilotEnrolled" `
        -Message "Autopilot enrolment failed" -Detail $errMsg
    Write-Warning "Autopilot enrolment failed: $_"
}

#=======================================================================
#   [OS] Collect Hardware Info
#   Moved to after Autopilot to avoid triggering driver reloads
#   during the install phase which caused unexpected reboots.
#=======================================================================

$cpuName  = (Get-CimInstance Win32_Processor).Name
$ram      = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
$diskSize = (Get-CimInstance Win32_DiskDrive | ForEach-Object { "{0} GB" -f ([math]::Round($_.Size / 1GB, 2)) }) -join ", "
$model    = (Get-CimInstance -ClassName Win32_ComputerSystem).Model

# Brief pause to allow MAC address to be readable
Start-Sleep -Seconds 5

# WiFi MAC - filters empty MACs and sorts so primary adapter wins
$wifiAdapter = Get-NetAdapter | Where-Object {
    ($_.Name -like "*Wi-Fi*" -or
     $_.InterfaceDescription -match "Wireless|Wi-Fi|802\.11|MediaTek|Intel.*WiFi|Qualcomm.*WiFi|Realtek.*WiFi") -and
    -not [string]::IsNullOrWhiteSpace($_.MacAddress)
} | Sort-Object Name | Select-Object -First 1

if ($wifiAdapter) {
    $wifiMac = ($wifiAdapter.MacAddress) -replace '-', ''
    Write-Host "Wi-Fi MAC : $wifiMac" -ForegroundColor Green
    Write-Host "Adapter   : $($wifiAdapter.Name) - $($wifiAdapter.InterfaceDescription)" -ForegroundColor Gray
} else {
    $wifiMac = "NOT_FOUND"
    Write-Host "ERROR: No Wi-Fi adapter found with a valid MAC address!" -ForegroundColor Red

    # Log all detected Wi-Fi adapters for diagnostics
    $allWifi = Get-NetAdapter | Where-Object {
        $_.Name -like "*Wi-Fi*" -or
        $_.InterfaceDescription -match "Wireless|Wi-Fi|802\.11"
    }

    $adapterDetail = ($allWifi | ForEach-Object {
        "$($_.Name) | MAC: '$($_.MacAddress)' | $($_.InterfaceDescription)"
    }) -join " || "

    Send-LogEvent -Level "ERROR" -Section "MACCollection" `
        -Message "No Wi-Fi adapter found with valid MAC address" `
        -Detail "Adapters detected: $adapterDetail"

    Write-Host "All Wi-Fi adapters detected:" -ForegroundColor Yellow
    $allWifi | ForEach-Object {
        Write-Host "  $($_.Name) | MAC: '$($_.MacAddress)' | $($_.InterfaceDescription)" -ForegroundColor Yellow
    }
}

Write-Host "CPU       : $cpuName"  -ForegroundColor Gray
Write-Host "RAM       : $ram GB"   -ForegroundColor Gray
Write-Host "Disk      : $diskSize" -ForegroundColor Gray
Write-Host "Model     : $model"    -ForegroundColor Gray

$hasMac = (-not [string]::IsNullOrWhiteSpace($wifiMac)) -and ($wifiMac -ne "NOT_FOUND")

#=======================================================================
#   [OS] Stage: Meraki Whitelist
#=======================================================================

Write-Host -ForegroundColor Green "Sending Meraki whitelist event"

if ($hasMac) {
    Send-BuildEvent -Stage "Meraki" -Extra @{ wifiMac = $wifiMac }
} else {
    Send-BuildEvent -Stage "Meraki" -Status "failed" `
        -ErrorMsg "No valid Wi-Fi MAC address - cannot whitelist device in Meraki"
    Send-LogEvent -Level "ERROR" -Section "Meraki" `
        -Message "Meraki whitelist skipped - no valid MAC address" `
        -Detail "wifiMac value: $wifiMac"
    Write-Warning "Meraki whitelist skipped - no valid MAC address available"
}

#=======================================================================
#   [OS] Stage: Jira Asset
#=======================================================================

Write-Host -ForegroundColor Green "Sending Jira asset creation event"

if ($hasMac) {
    Send-BuildEvent -Stage "JiraAsset" -Status "success" -Extra @{
        cpuName  = $cpuName
        ram      = "$ram GB"
        diskSize = $diskSize
        model    = $model
        wifiMac  = $wifiMac
    }
} else {
    Send-BuildEvent -Stage "JiraAsset" -Status "failed" `
        -ErrorMsg "No valid Wi-Fi MAC address - Jira asset created without MAC"
    Send-LogEvent -Level "ERROR" -Section "JiraAsset" `
        -Message "JiraAsset sent with failure - no valid MAC address" `
        -Detail "wifiMac value: $wifiMac"
    Write-Warning "JiraAsset event sent with failure status - no valid MAC address"
}

#=======================================================================
#   [OS] Stage: BuildComplete
#=======================================================================

Send-BuildEvent -Stage "BuildComplete"

#=======================================================================
#   [OS] Tidy Up
#=======================================================================

Remove-Item -Path "C:\Temp"        -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\OSDCloud\"   -Recurse -Force -ErrorAction SilentlyContinue
