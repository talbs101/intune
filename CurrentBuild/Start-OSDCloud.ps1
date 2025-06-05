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

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

#-----------------------------------------------
#   Ask for Device Name
#-----------------------------------------------
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

#-----------------------------------------------
#   Ask for Build Type & Builder via WinForms
#-----------------------------------------------
$buildType   = $null
$builder     = $null
$form        = New-Object System.Windows.Forms.Form
$form.Text   = "Build Selector"
$form.Size   = New-Object System.Drawing.Size(350,300)
$form.StartPosition = "CenterScreen"

# Device Name Label
$labelDevice = New-Object System.Windows.Forms.Label
$labelDevice.Text     = "Device: $deviceName"
$labelDevice.AutoSize = $true
$labelDevice.Location = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($labelDevice)

# Label: What build type?
$labelBuild = New-Object System.Windows.Forms.Label
$labelBuild.Text     = "What build type?"
$labelBuild.AutoSize = $true
$labelBuild.Location = New-Object System.Drawing.Point(20,60)
$form.Controls.Add($labelBuild)

# ComboBox: Build Type
$comboBuild = New-Object System.Windows.Forms.ComboBox
$comboBuild.Location    = New-Object System.Drawing.Point(20,90)
$comboBuild.Size        = New-Object System.Drawing.Size(280,24)
$comboBuild.DropDownStyle = 'DropDownList'
$comboBuild.Items.AddRange(@("Standard","Shared","Kiosk","Windows 11"))
$form.Controls.Add($comboBuild)

# Label: Who is building?
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
$button.Location = New-Object System.Drawing.Point(20,210)
$button.Size     = New-Object System.Drawing.Size(280,30)
$button.Text     = "Start"
$button.Add_Click({
    $selectedBuild   = $comboBuild.SelectedItem
    $selectedBuilder = $comboBuilder.SelectedItem

    if (-not $selectedBuild -or -not $selectedBuilder) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select both a build type and a builder.",
            "Selection Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Assign to outer‐scope variables and close form
    $script:buildType = $selectedBuild
    $script:builder   = $selectedBuilder
    $form.Close()
})
$form.Controls.Add($button)

# Show the form (modal) and wait for user input
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

# If user closed the form without clicking Start, exit
if (-not $buildType -or -not $builder) {
    Write-Host "No build type or builder selected. Exiting."
    exit
}

#================================================
#   [PreOS] Hypervisor‐Specific Configuration
#================================================
if ((Get-MyComputerModel) -match 'Virtual') {
    Write-Host -ForegroundColor Green "Setting Display Resolution to 1600x900"
    Set-DisRes 1600
}

#================================================
#   [PreOS] Import OSD Module
#================================================
Write-Host -ForegroundColor Green "Importing OSD PowerShell Module"
Import-Module OSD -Force

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

Write-Host -ForegroundColor Green "Writing Device Name to C:\OSDCloud\DeviceName.txt"
Set-Content -Path "C:\OSDCloud\DeviceName.txt" -Value $deviceName -Force

#================================================
#  [PostOS] Autopilot OOBE CMD (keep this as-is)
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
#  [PostOS] SetupComplete CMD – Varies By Build Type
#================================================
# Define which GitHub script to invoke based on the build type:
switch ($buildType) {
    "Standard" {
        $scriptUrl = "https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/Standard.ps1"
        Set-Content -Path "C:\OSDCloud\BuildType.txt" -Value $buildType -Force
        Set-Content -Path "C:\OSDCloud\Builder.txt" -Value $builder -Force
    }
    "Shared" {
        $scriptUrl = "https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/Standard.ps1"
        Set-Content -Path "C:\OSDCloud\BuildType.txt" -Value $buildType -Force
        Set-Content -Path "C:\OSDCloud\Builder.txt" -Value $builder -Force
    }
    "Kiosk" {
        $scriptUrl = "https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/Standard.ps1"
        Set-Content -Path "C:\OSDCloud\BuildType.txt" -Value $buildType -Force
        Set-Content -Path "C:\OSDCloud\Builder.txt" -Value $builder -Force
    }
    "Windows 11" {
        $scriptUrl = "https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/Standard.ps1"
        Set-Content -Path "C:\OSDCloud\BuildType.txt" -Value $buildType -Force
        Set-Content -Path "C:\OSDCloud\Builder.txt" -Value $builder -Force
    }
    default {
        # Fallback in case something unexpected happened
        $scriptUrl = "https://raw.githubusercontent.com/talbs101/intune/refs/heads/main/CurrentBuild/Standard.ps1"
    }
}

Write-Host -ForegroundColor Green "Creating C:\Windows\Setup\Scripts\SetupComplete.cmd (BuildType = $buildType)"
$SetupCompleteCMD = @"
powershell.exe -Command Set-ExecutionPolicy RemoteSigned -Force
powershell.exe -Command "& {IEX (IRM $scriptUrl)}"
"@
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
