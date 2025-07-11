#================================================
#   [PreOS] Update Module
#================================================

#Write-Host -ForegroundColor Green "Updating OSD PowerShell Module"
#Install-Module OSD -Force

Write-Host  -ForegroundColor Green "Importing OSD PowerShell Module"
Import-Module OSD -Force 

#================================================
#   [PreOS] Update Module & Environment Preparation
#================================================

Add-Type -AssemblyName System.Windows.Forms

# ───────────────────────────────────────────────────────────────────────
# Step 1: Ask for Device Name via VB InputBox (unchanged)
# ───────────────────────────────────────────────────────────────────────
Add-Type -AssemblyName Microsoft.VisualBasic
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
# Step 2: Build‐Type / Builder Form
# ───────────────────────────────────────────────────────────────────────
# Prepare variables to hold user selections
$buildType = $null
$builder   = $null

$form = New-Object System.Windows.Forms.Form
$form.Text          = "Build Selector"
$form.Size          = New-Object System.Drawing.Size(350,300)
$form.StartPosition = "CenterScreen"

# Label: show the device name
$labelDevice = New-Object System.Windows.Forms.Label
$labelDevice.Text     = "Device: $deviceName"
$labelDevice.AutoSize = $true
$labelDevice.Location = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($labelDevice)

# Label: "What build type?"
$labelBuild = New-Object System.Windows.Forms.Label
$labelBuild.Text     = "What build type?"
$labelBuild.AutoSize = $true
$labelBuild.Location = New-Object System.Drawing.Point(20,60)
$form.Controls.Add($labelBuild)

# ComboBox: Build Type
$comboBuild = New-Object System.Windows.Forms.ComboBox
$comboBuild.Location     = New-Object System.Drawing.Point(20,90)
$comboBuild.Size         = New-Object System.Drawing.Size(280,24)
$comboBuild.DropDownStyle = 'DropDownList'
$comboBuild.Items.AddRange(@("Standard","Shared","Kiosk","Windows 11", "Rebuild-Standard", "Rebuild-Shared"))
$form.Controls.Add($comboBuild)

# Label: "Who is building?"
$labelBuilder = New-Object System.Windows.Forms.Label
$labelBuilder.Text     = "Who is building?"
$labelBuilder.AutoSize = $true
$labelBuilder.Location = New-Object System.Drawing.Point(20,130)
$form.Controls.Add($labelBuilder)

# ComboBox: Builder
$comboBuilder = New-Object System.Windows.Forms.ComboBox
$comboBuilder.Location     = New-Object System.Drawing.Point(20,160)
$comboBuilder.Size         = New-Object System.Drawing.Size(280,24)
$comboBuilder.DropDownStyle = 'DropDownList'
$comboBuilder.Items.AddRange(@("Jon","Ryan","Jacob","Mike","James T"))
$form.Controls.Add($comboBuilder)

# Button: Start
$button = New-Object System.Windows.Forms.Button
$button.Location      = New-Object System.Drawing.Point(20,210)
$button.Size          = New-Object System.Drawing.Size(280,30)
$button.Text          = "Start"
# **This is critical**: set the button’s DialogResult to OK so ShowDialog() will return OK
$button.DialogResult  = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($button)

# Make this the “accept” button (activated by pressing Enter)
$form.AcceptButton = $button

# Show the form modally
$result = $form.ShowDialog()

# If the user clicked “X” or closed the window in any other way, quit
if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "User cancelled or closed the form. Exiting."
    exit
}

# Now that ShowDialog() returned OK, read the selected items
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

# At this point, $buildType and $builder are guaranteed to be non‐null
Write-Host "Selected Build Type: $buildType"
Write-Host "Selected Builder:    $builder"

# ───────────────────────────────────────────────────────────────────────
# Step 3: (Example) Continue with Start-OSDCloud or whatever comes next
# ───────────────────────────────────────────────────────────────────────
# For demonstration, we’ll just print them. In your actual script, you’d
# pass these variables to Start-OSDCloud or write out your SetupComplete.cmd, etc.

# Example placeholder:
Write-Host "Now calling Start-OSDCloud with $buildType by $builder ..."
# Start-OSDCloud @Params  # ← your real logic goes here


#================================================
#   [PreOS] Hypervisor‐Specific Configuration
#================================================
if ((Get-MyComputerModel) -match 'Virtual') {
    Write-Host -ForegroundColor Green "Setting Display Resolution to 1600x900"
    Set-DisRes 1600
}

#================================================
#   [OS] Start OSDCloud
#================================================
$Params = @{
    OSVersion   = "Windows 11"
    OSBuild     = "24H2"
    OSEdition   = "Pro"
    OSLanguage  = "en-gb"
    OSLicense   = "Volume"
    ZTI         = $true
    Firmware    = $false
}
Write-Host -ForegroundColor Green "Starting OSDCloud (OSVersion=$($Params.OSVersion), Build=$($Params.OSBuild))"
Start-OSDCloud @Params

#================================================
#  [PostOS] OOBEDeploy Configuration
#================================================
Write-Host -ForegroundColor Green "Copying SetupComplete dependencies..."
Copy-Item "X:\OSDCloud\Config\Scripts\SetupComplete\Secrets.ps1" "C:\OSDCloud\Scripts\Secrets.ps1" -Force
Copy-Item "X:\OSDCloud\Config\Scripts\SetupComplete\Get-WindowsAutoPilotInfo.ps1" "C:\OSDCloud\Scripts\Get-WindowsAutoPilotInfo.ps1" -Force

# Copy CompanyPortal files

# source & destination roots
$sourceRoot      = "X:\OSDCloud\Config\Scripts\SetupComplete\Apps\CompanyPortal"
$destinationRoot = "C:\OSDCloud\CompanyPortal"

# all files, *including* their relative paths under Dependencies
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
    # build full paths
    $src  = Join-Path $sourceRoot      $rel
    $dest = Join-Path $destinationRoot $rel

    # ensure destination folder exists
    $destDir = Split-Path $dest -Parent
    if (-not (Test-Path $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
    }

    # copy
    Write-Host "Copying `"$src`" → `"$dest`""
    Copy-Item -Path $src -Destination $dest -Force
}

Write-Host "All files copied." 


Set-Content -Path "C:\OSDCloud\DeviceName.txt" -Value $deviceName -Force
Set-Content -Path "C:\OSDCloud\BuildType.txt" -Value $buildType -Force
Set-Content -Path "C:\OSDCloud\Builder.txt" -Value $builder -Force

#================================================
#  [PostOS] OOBEDeploy Configuration
#================================================

# Copy secrets into Secrets.ps1 and store on C:
Copy-Item "X:\OSDCloud\Config\Scripts\SetupComplete\Secrets.ps1" "C:\OSDCloud\Scripts\Secrets.ps1" -Force

Write-Host -ForegroundColor Green "Writing Device Name to C:\OSDCloud\DeviceName.txt"
Set-Content -Path "C:\OSDCloud\DeviceName.txt" -Value $deviceName

#================================================
#  [PostOS] Autopilot OOBE CMD (keep this as-is)
#================================================
Write-Host -ForegroundColor Green "Creating C:\Windows\System32\OOBE.cmd"
$OOBECMD = @'
PowerShell -NoL -Com Set-ExecutionPolicy RemoteSigned -Force
Set Path = %PATH%;C:\Program Files\WindowsPowerShell\Scripts
Start /Wait PowerShell -NoL -C Install-Module AutopilotOOBE -Force -Verbose
Start /Wait PowerShell -NoL -C Install-Module OSD -Force -Verbose
Start /Wait PowerShell -NoL -C Start-OOBEDeploy
Start /Wait PowerShell -NoL -C Invoke-WebPSScript https://tpm.osdcloud.ch
Start /Wait PowerShell -NoL -C Invoke-WebPSScript https://cleanup.osdcloud.ch
Start /Wait PowerShell -NoL -C Restart-Computer -Force
'@
$OOBECMD | Out-File -FilePath 'C:\Windows\System32\OOBE.cmd' -Encoding ascii -Force


#================================================
#  [PostOS] SetupComplete CMD Command Line
#================================================
Write-Host -ForegroundColor Green "Create C:\Windows\Setup\Scripts\SetupComplete.cmd"
$SetupCompleteCMD = @'
powershell.exe -Command Set-ExecutionPolicy RemoteSigned -Force
powershell.exe -Command "& {IEX (IRM https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/CompanyPortal.ps1)}"
powershell.exe -Command "& {IEX (IRM https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/NewBuild/Standard.ps1)}"
'@
$SetupCompleteCMD | Out-File -FilePath 'C:\Windows\Setup\Scripts\SetupComplete.cmd' -Encoding ascii -Force
# Ensure the Setup Scripts folder exists:
if (!(Test-Path "C:\Windows\Setup\Scripts")) {
    New-Item "C:\Windows\Setup\Scripts" -ItemType Directory -Force | Out-Null
}
$SetupCompleteCMD | Out-File -FilePath 'C:\Windows\Setup\Scripts\SetupComplete.cmd' -Encoding ascii -Force

#================================================
#   [PostOS] Restart to complete task sequence
#================================================
Write-Host -ForegroundColor Green "Restarting in 20 seconds..."
Start-Sleep -Seconds 20
wpeutil reboot
