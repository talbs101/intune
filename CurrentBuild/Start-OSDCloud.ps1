#================================================
#   [PreOS] Update Module & Environment Preparation
#================================================

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

#-----------------------------------------------
#   Ask for Device Name via VB InputBox
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
$buildType = $null
$builder   = $null

$form = New-Object System.Windows.Forms.Form
$form.Text          = "Build Selector"
$form.Size          = New-Object System.Drawing.Size(350,300)
$form.StartPosition = "CenterScreen"

# Label: Device
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
$comboBuild.Location     = New-Object System.Drawing.Point(20,90)
$comboBuild.Size         = New-Object System.Drawing.Size(280,24)
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
    $selBuild   = $comboBuild.SelectedItem
    $selBuilder = $comboBuilder.SelectedItem

    if (-not $selBuild -or -not $selBuilder) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select both a build type and a builder.",
            "Selection Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Store into script‐scope variables and close form
    $script:buildType = $selBuild
    $script:builder   = $selBuilder
    $form.Close()
})
$form.Controls.Add($button)

# Show the form and wait
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

# If the user closed the window without clicking Start, exit now
if (-not $buildType -or -not $builder) {
    Write-Host "No build type or builder selected. Exiting."
    exit
}
