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
Add-Type -AssemblyName System.Drawing

$deviceName = $env:COMPUTERNAME
$buildType  = $null
$builder    = $null

#================================================
#   [PreOS] Build Selector Form
#   No nested function definitions or Add_Click scriptblocks
#   — IEX/Invoke-WebPSScript compatible
#================================================

$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "Laptop Build"
$form.Size             = New-Object System.Drawing.Size(420, 420)
$form.StartPosition    = "CenterScreen"
$form.BackColor        = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.FormBorderStyle  = "FixedDialog"
$form.MaximizeBox      = $false
$form.MinimizeBox      = $false
$form.Font             = New-Object System.Drawing.Font("Segoe UI", 9)

# ── Header bar ──────────────────────────────────────────────
$header              = New-Object System.Windows.Forms.Panel
$header.Size         = New-Object System.Drawing.Size(420, 60)
$header.Location     = New-Object System.Drawing.Point(0, 0)
$header.BackColor    = [System.Drawing.Color]::FromArgb(17, 24, 39)
$form.Controls.Add($header)

$headerLabel           = New-Object System.Windows.Forms.Label
$headerLabel.Text      = "Laptop Build — St Monica Trust IT"
$headerLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$headerLabel.ForeColor = [System.Drawing.Color]::White
$headerLabel.AutoSize  = $true
$headerLabel.Location  = New-Object System.Drawing.Point(20, 10)
$header.Controls.Add($headerLabel)

$subLabel              = New-Object System.Windows.Forms.Label
$subLabel.Text         = "Complete all fields before starting the build"
$subLabel.Font         = New-Object System.Drawing.Font("Segoe UI", 8)
$subLabel.ForeColor    = [System.Drawing.Color]::FromArgb(156, 163, 175)
$subLabel.AutoSize     = $true
$subLabel.Location     = New-Object System.Drawing.Point(20, 34)
$header.Controls.Add($subLabel)

# ── Serial number display ────────────────────────────────────
$serialPanel             = New-Object System.Windows.Forms.Panel
$serialPanel.Size        = New-Object System.Drawing.Size(378, 36)
$serialPanel.Location    = New-Object System.Drawing.Point(20, 76)
$serialPanel.BackColor   = [System.Drawing.Color]::FromArgb(239, 246, 255)
$serialPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($serialPanel)

$serialLabel           = New-Object System.Windows.Forms.Label
$serialLabel.Text      = "Serial Number"
$serialLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
$serialLabel.ForeColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
$serialLabel.AutoSize  = $true
$serialLabel.Location  = New-Object System.Drawing.Point(10, 4)
$serialPanel.Controls.Add($serialLabel)

$serialValue           = New-Object System.Windows.Forms.Label
$serialValue.Text      = $serial
$serialValue.Font      = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$serialValue.ForeColor = [System.Drawing.Color]::FromArgb(17, 24, 39)
$serialValue.AutoSize  = $true
$serialValue.Location  = New-Object System.Drawing.Point(10, 18)
$serialPanel.Controls.Add($serialValue)

# ── Device Name label ────────────────────────────────────────
$lblDevice           = New-Object System.Windows.Forms.Label
$lblDevice.Text      = "Device Name"
$lblDevice.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblDevice.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
$lblDevice.AutoSize  = $true
$lblDevice.Location  = New-Object System.Drawing.Point(20, 130)
$form.Controls.Add($lblDevice)

# ── Device Name input ────────────────────────────────────────
$txtDevice             = New-Object System.Windows.Forms.TextBox
$txtDevice.Location    = New-Object System.Drawing.Point(20, 148)
$txtDevice.Size        = New-Object System.Drawing.Size(378, 28)
$txtDevice.Font        = New-Object System.Drawing.Font("Segoe UI", 10)
$txtDevice.Text        = $env:COMPUTERNAME
$txtDevice.BackColor   = [System.Drawing.Color]::White
$txtDevice.BorderStyle = "FixedSingle"
$form.Controls.Add($txtDevice)

# ── Build Type label ─────────────────────────────────────────
$lblBuild            = New-Object System.Windows.Forms.Label
$lblBuild.Text       = "Build Type"
$lblBuild.Font       = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblBuild.ForeColor  = [System.Drawing.Color]::FromArgb(55, 65, 81)
$lblBuild.AutoSize   = $true
$lblBuild.Location   = New-Object System.Drawing.Point(20, 192)
$form.Controls.Add($lblBuild)

# ── Build Type dropdown ──────────────────────────────────────
$comboBuild               = New-Object System.Windows.Forms.ComboBox
$comboBuild.Location      = New-Object System.Drawing.Point(20, 210)
$comboBuild.Size          = New-Object System.Drawing.Size(378, 28)
$comboBuild.Font          = New-Object System.Drawing.Font("Segoe UI", 10)
$comboBuild.DropDownStyle = "DropDownList"
$comboBuild.BackColor     = [System.Drawing.Color]::White
$comboBuild.Items.AddRange(@("Standard", "Care", "Kiosk-Chapel"))
$form.Controls.Add($comboBuild)

# ── Builder label ────────────────────────────────────────────
$lblBuilder           = New-Object System.Windows.Forms.Label
$lblBuilder.Text      = "Builder"
$lblBuilder.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblBuilder.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
$lblBuilder.AutoSize  = $true
$lblBuilder.Location  = New-Object System.Drawing.Point(20, 254)
$form.Controls.Add($lblBuilder)

# ── Builder dropdown ─────────────────────────────────────────
$comboBuilder               = New-Object System.Windows.Forms.ComboBox
$comboBuilder.Location      = New-Object System.Drawing.Point(20, 272)
$comboBuilder.Size          = New-Object System.Drawing.Size(378, 28)
$comboBuilder.Font          = New-Object System.Drawing.Font("Segoe UI", 10)
$comboBuilder.DropDownStyle = "DropDownList"
$comboBuilder.BackColor     = [System.Drawing.Color]::White
$comboBuilder.Items.AddRange(@(
    "jon.titchmarsh@stmonicatrust.org.uk",
    "ryan.appleton@stmonicatrust.org.uk",
    "jacob.ashby@stmonicatrust.org.uk",
    "mike.grunshaw@stmonicatrust.org.uk",
    "james.talbot@stmonicatrust.org.uk"
))
$form.Controls.Add($comboBuilder)

# ── Start Build button ───────────────────────────────────────
$btnStart                           = New-Object System.Windows.Forms.Button
$btnStart.Location                  = New-Object System.Drawing.Point(20, 326)
$btnStart.Size                      = New-Object System.Drawing.Size(378, 44)
$btnStart.Text                      = "START BUILD"
$btnStart.Font                      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnStart.BackColor                 = [System.Drawing.Color]::FromArgb(22, 163, 74)
$btnStart.ForeColor                 = [System.Drawing.Color]::White
$btnStart.FlatStyle                 = "Flat"
$btnStart.FlatAppearance.BorderSize = 0
$btnStart.DialogResult              = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($btnStart)
$form.AcceptButton                  = $btnStart

# ── Show form ────────────────────────────────────────────────
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "User cancelled. Exiting."
    exit
}

# ── Read values ──────────────────────────────────────────────
$deviceName = $txtDevice.Text.Trim()
$buildType  = $comboBuild.SelectedItem
$builder    = $comboBuilder.SelectedItem

# ── Validate after form closes ───────────────────────────────
if ([string]::IsNullOrWhiteSpace($deviceName)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Please enter a device name.",
        "Required", "OK", "Warning")
    exit
}

if (-not $buildType) {
    [System.Windows.Forms.MessageBox]::Show(
        "Please select a build type.",
        "Required", "OK", "Warning")
    exit
}

if (-not $builder) {
    [System.Windows.Forms.MessageBox]::Show(
        "Please select a builder.",
        "Required", "OK", "Warning")
    exit
}

Write-Host "Device    : $deviceName" -ForegroundColor Green
Write-Host "Build Type: $buildType"  -ForegroundColor Green
Write-Host "Builder   : $builder"    -ForegroundColor Green

#=======================================================================
#   [OS] Send OSDStarted Event
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
    OSVersion  = "Windows 11"
    OSBuild    = "24H2"
    OSEdition  = "Pro"
    OSLanguage = "en-gb"
    OSLicense  = "Volume"
    ZTI        = $true
    Firmware   = $false
}

Write-Host -ForegroundColor Green "Starting OSDCloud (OSVersion=$($Params.OSVersion), Build=$($Params.OSBuild))"
Start-OSDCloud @Params

#=======================================================================
#   [OS] Send OSDComplete Event
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
    Write-Host "Copying $src to $dest"
    Copy-Item -Path $src -Destination $dest -Force
}

Write-Host "All files copied."

#================================================
#   [PostOS] Write Build Info Files
#================================================

Set-Content -Path "C:\OSDCloud\DeviceName.txt" -Value $deviceName -Force
Set-Content -Path "C:\OSDCloud\BuildType.txt"  -Value $buildType  -Force
Set-Content -Path "C:\OSDCloud\Builder.txt"    -Value $builder    -Force

Write-Host -ForegroundColor Green "Writing build info files to C:\OSDCloud\"

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
powershell.exe -Command "& {IEX (IRM https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/Standard-v2.ps1)}"
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
