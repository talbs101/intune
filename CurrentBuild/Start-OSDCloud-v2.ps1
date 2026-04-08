#================================================
#   [PreOS] Update Module
#================================================

Write-Host -ForegroundColor Green "Importing OSD PowerShell Module"
Import-Module OSD -Force

#================================================
#   [PreOS] Load Secrets
#   Must be early so $LogicAppUrl is available for Send-BuildEvent
#================================================

$secretsFile = 'X:\OSDCloud\Config\Scripts\SetupComplete\Secrets.ps1'
if (Test-Path $secretsFile) {
    . $secretsFile
} else {
    Write-Warning "Secrets file not found at $secretsFile — build events will not fire"
}

$LogicAppUrl = $env:BUILD_LogicAppUrl2

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

        Write-Host "[$Stage] Event sent — $Status" -ForegroundColor Cyan
        return $response

    } catch {
        Write-Warning "[$Stage] Failed to send build event: $_"
        return $null
    }
}

#================================================
#   [PreOS] Collect Serial Number
#   Available in WinPE — needed for Table Storage key and Jira asset
#================================================

$serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
Write-Host -ForegroundColor Gray "Serial Number: $serial"

#================================================
#   [PreOS] Environment Preparation
#================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# Initialise so Send-BuildEvent doesn't error if form is cancelled
$deviceName = $env:COMPUTERNAME
$buildType  = $null
$builder    = $null

# ───────────────────────────────────────────────────────────────────────
# Step 1: Ask for Device Name
# ───────────────────────────────────────────────────────────────────────

$deviceName = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the device name:",
    "Device Name Required",
    "$env:COMPUTERNAME"
)

if ([string]::IsNullOrWhiteSpace($deviceName)) {
    [System.Windows.Forms.MessageBox]::Show(
        "You must enter a device name to continue.",
        "Device Name Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    exit
}

# ───────────────────────────────────────────────────────────────────────
# Step 2: Build Type / Builder Form
# ───────────────────────────────────────────────────────────────────────

$form = New-Object System.Windows.Forms.Form
$form.Text          = "Build Selector"
$form.Size          = New-Object System.Drawing.Size(350,300)
$form.StartPosition = "CenterScreen"

$labelDevice           = New-Object System.Windows.Forms.Label
$labelDevice.Text      = "Device: $deviceName"
$labelDevice.AutoSize  = $true
$labelDevice.Location  = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($labelDevice)

$labelBuild            = New-Object System.Windows.Forms.Label
$labelBuild.Text       = "What build type?"
$labelBuild.AutoSize   = $true
$labelBuild.Location   = New-Object System.Drawing.Point(20,60)
$form.Controls.Add($labelBuild)

$comboBuild                = New-Object System.Windows.Forms.ComboBox
$comboBuild.Location       = New-Object System.Drawing.Point(20,90)
$comboBuild.Size           = New-Object System.Drawing.Size(280,24)
$comboBuild.DropDownStyle  = 'DropDownList'
$comboBuild.Items.AddRange(@("Standard", "Care", "Kiosk-Chapel"))
$form.Controls.Add($comboBuild)

$labelBuilder          = New-Object System.Windows.Forms.Label
$labelBuilder.Text     = "Who is building?"
$labelBuilder.AutoSize = $true
$labelBuilder.Location = New-Object System.Drawing.Point(20,130)
$form.Controls.Add($labelBuilder)

$comboBuilder              = New-Object System.Windows.Forms.ComboBox
$comboBuilder.Location     = New-Object System.Drawing.Point(20,160)
$comboBuilder.Size         = New-Object System.Drawing.Size(280,24)
$comboBuilder.DropDownStyle = 'DropDownList'
$comboBuilder.Items.AddRange(@(
    "jon.titchmarsh@stmonicatrust.org.uk",
    "ryan.appleton@stmonicatrust.org.uk",
    "jacob.ashby@stmonicatrust.org.uk",
    "mike.grunshaw@stmonicatrust.org.uk",
    "james.talbot@stmonicatrust.org.uk"
))
$form.Controls.Add($comboBuilder)

$button                = New-Object System.Windows.Forms.Button
$button.Location       = New-Object System.Drawing.Point(20,210)
$button.Size           = New-Object System.Drawing.Size(280,30)
$button.Text           = "Start"
$button.DialogResult   = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($button)
$form.AcceptButton     = $button

$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "User cancelled. Exiting."
    exit
}

$buildType = $comboBuild.SelectedItem
$builder   = $comboBuilder.SelectedItem

if (-not $buildType -or -not $builder) {
    [System.Windows.Forms.MessageBox]::Show(
        "Please select both a build type and a builder.",
        "Selection Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    exit
}

Write-Host "Device    : $deviceName" -ForegroundColor Green
Write-Host "Build Type: $buildType"  -ForegroundColor Green
Write-Host "Builder   : $builder"    -ForegroundColor Green

#=======================================================================
#   [OS] Send OSDStarted Event
#   Fires as soon as helpdesk has confirmed device name, build type
#   and builder — before imaging begins
#=======================================================================

Send-BuildEvent -Stage "OSDStarted" -Extra @{
    user = $builder
}

#================================================
#   [PreOS] Hypervisor-Specific Configuration
#================================================

if ((Get-MyComputerModel) -match 'Virtual') {
    Write-Host -ForegroundColor Green "Setting Display Resolution to 1600x900"
    Set-DisRes 1600
}

#================================================
#   [OS] Start OSDCloud
#================================================

$Params = @{
    OSVersion = "Windows 11"
    OSBuild   = "24H2"
    OSEdition = "Pro"
    OSLanguage = "en-gb"
    OSLicense = "Volume"
    ZTI       = $true
    Firmware  = $false
}

Write-Host -ForegroundColor Green "Starting OSDCloud (OSVersion=$($Params.OSVersion), Build=$($Params.OSBuild))"
Start-OSDCloud @Params

#=======================================================================
#   [OS] Send OSDComplete Event
#   Fires immediately after Start-OSDCloud returns — imaging is done,
#   SetupComplete scripts will run on next boot
#=======================================================================

Send-BuildEvent -Stage "OSDComplete"

#================================================
#   [PostOS] Copy SetupComplete Dependencies
#================================================

Write-Host -ForegroundColor Green "Copying SetupComplete dependencies..."
Copy-Item "X:\OSDCloud\Config\Scripts\SetupComplete\Secrets.ps1" "C:\OSDCloud\Scripts\Secrets.ps1" -Force
Copy-Item "X:\OSDCloud\Config\Scripts\SetupComplete\Get-WindowsAutoPilotInfo.ps1" "C:\OSDCloud\Scripts\Get-WindowsAutoPilotInfo.ps1" -Force

# Copy CompanyPortal files
$sourceRoot      = "X:\OSDCloud\Config\Scripts\SetupComplete\Apps\CompanyPortal"
$destinationRoot = "C:\OSDCloud\CompanyPortal"

$filesToCopy = @(
    "CompanyPortal.appxbundle",
    "Dependencies\AUMIDs.txt",
    "Dependencies\c797dbb4414543f59d35e59e5225824e_License1.xml",
    "Dependencies\MPAP_c797dbb4414543f59d35e59e5225824e_001.provxml",
    "Dependencies\Microsoft.NET.Native.Framework.2.2_2.2.29512.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies\Microsoft.NET.Native.Runtime.2.2_2.2.28604.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies\Microsoft.Services.Store.Engagement_10.0.23012.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies\Microsoft.UI.Xaml.2.7_7.2409.9001.0_x64__8wekyb3d8bbwe.appx",
    "Dependencies\Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.appx"
)

foreach ($rel in $filesToCopy) {
    $src     = Join-Path $sourceRoot $rel
    $dest    = Join-Path $destinationRoot $rel
    $destDir = Split-Path $dest -Parent
    if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
    Write-Host "Copying `"$src`" → `"$dest`""
    Copy-Item -Path $src -Destination $dest -Force
}

Write-Host "All files copied."

#================================================
#   [PostOS] Write Build Info Files
#   These are read by SetupComplete.ps1 on first boot
#================================================

Set-Content -Path "C:\OSDCloud\DeviceName.txt" -Value $deviceName -Force
Set-Content -Path "C:\OSDCloud\BuildType.txt"  -Value $buildType  -Force
Set-Content -Path "C:\OSDCloud\Builder.txt"    -Value $builder    -Force

Write-Host -ForegroundColor Green "Writing Device Name to C:\OSDCloud\DeviceName.txt"

#================================================
#   [PostOS] Secrets
#================================================

Copy-Item "X:\OSDCloud\Config\Scripts\SetupComplete\Secrets.ps1" "C:\OSDCloud\Scripts\Secrets.ps1" -Force

#================================================
#   [PostOS] OOBE CMD
#================================================

Write-Host -ForegroundColor Green "Creating C:\Windows\System32\OOBE.cmd"
$OOBECMD = @'
PowerShell -NoL -Com Set-ExecutionPolicy RemoteSigned -Force
Set Path = %PATH%;C:\Program Files\WindowsPowerShell\Scripts
Start /Wait PowerShell -NoL -C Install-Module AutopilotOOBE -Force -Verbose
Start /Wait PowerShell -NoL -C Install-Module OSD -Force -Verbose
Start /Wait PowerShell -NoL -C Invoke-WebPSScript https://check-autopilotprereq.osdcloud.ch
Start /Wait PowerShell -NoL -C Start-OOBEDeploy
Start /Wait PowerShell -NoL -C Invoke-WebPSScript https://tpm.osdcloud.ch
Start /Wait PowerShell -NoL -C Invoke-WebPSScript https://cleanup.osdcloud.ch
Start /Wait PowerShell -NoL -C Restart-Computer -Force
'@
$OOBECMD | Out-File -FilePath 'C:\Windows\System32\OOBE.cmd' -Encoding ascii -Force

#================================================
#   [PostOS] SetupComplete CMD
#================================================

Write-Host -ForegroundColor Green "Create C:\Windows\Setup\Scripts\SetupComplete.cmd"
$SetupCompleteCMD = @'
powershell.exe -Command Set-ExecutionPolicy RemoteSigned -Force
powershell.exe -Command "& {IEX (IRM https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/CompanyPortal.ps1)}"
powershell.exe -Command "& {IEX (IRM https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/Standard.ps1)}"
'@

if (!(Test-Path "C:\Windows\Setup\Scripts")) {
    New-Item "C:\Windows\Setup\Scripts" -ItemType Directory -Force | Out-Null
}
$SetupCompleteCMD | Out-File -FilePath 'C:\Windows\Setup\Scripts\SetupComplete.cmd' -Encoding ascii -Force

#================================================
#   [PostOS] Restart
#================================================

Write-Host -ForegroundColor Green "Restarting in 20 seconds..."
Start-Sleep -Seconds 20
wpeutil reboot
