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
$Office2019Url      = $env:BUILD_Office2019Url
$Office2019XMLUrl   = $env:BUILD_Office2019XMLUrl


$DeviceNameFile  = "C:\OSDCloud\DeviceName.txt"
$BuildTypeFile   = "C:\OSDCloud\BuildType.txt"
$BuilderFile     = "C:\OSDCloud\Builder.txt"


#=======================================================================
#   [OS] Install Office 365
#=======================================================================

<# ---------------------------------------------------------------------------
 Office installer – chooses Office 365 or Office 2019 based on BuildType.txt
    • Shared    → download + install Office 2019
    • Standard  → install Office 365
    • Rebuild   → install Office 365
    • Anything else → default to Office 365
---------------------------------------------------------------------------#>

# 1. URLs – adjust to your storage locations

# 2. Local working paths
$workingDir     = 'C:\Temp'
$localSetupPath = Join-Path $workingDir 'setup.exe'
$localXmlPath   = Join-Path $workingDir 'install.xml'

# 3. Ensure TLS 1.2 (needed in WinPE/WinRE for Azure blobs)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 4. Read build type, trimming CR/LF and possible UTF‑8 BOM
$buildTypeFile = 'C:\OSDCloud\BuildType.txt'
$buildType     = (Get-Content $buildTypeFile -Raw).Trim().Trim([char]0xFEFF)

# 5. Decide which package to use and whether we must pre‑download
$needsPreDownload = $false
switch -Regex ($buildType) {
    '^(?i)shared$' {
        $blobUrl          = $Office2019Url
        $xmlUrl           = $Office2019XMLUrl
        $message          = 'Install Office 2019 for Shared Machine'
        $needsPreDownload = $true           # Shared machines: stage source first
    }
    '^(?i)(standard|rebuild)$' {
        $blobUrl = $Office365Url
        $xmlUrl  = $Office365XMLUrl
        $message = 'Installing Office 365'
    }
    default {
        $blobUrl = $Office365Url
        $xmlUrl  = $Office365XMLUrl
        $message = "Can't determine build type, installing 365"
    }
}

Write-Host -ForegroundColor Cyan  "Build type detected: $buildType"
Write-Host -ForegroundColor Green $message

# 6. Ensure working directory exists
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
}

# 7. Download setup.exe and install.xml (‑UseBasicParsing avoids IE parser)
Invoke-WebRequest -Uri $blobUrl -OutFile $localSetupPath -UseBasicParsing
Invoke-WebRequest -Uri $xmlUrl  -OutFile $localXmlPath  -UseBasicParsing

# 8. Pre‑download source files only for Shared (Office 2019) machines
Write-Host -ForegroundColor Yellow 'Pre‑downloading Office source files …'
$installArgs1 = @("/download configuration.xml")
Start-Process -FilePath $localSetupPath -ArgumentList $installArgs1 -Wait -NoNewWindow

# 9. Install Office
#$installArgs2 = @("/configure `"$localXmlPath`"")
#Start-Process -FilePath $localSetupPath -ArgumentList $installArgs1 -Wait -NoNewWindow




#=======================================================================
#   [OS] Install Company Portal
#=======================================================================

Write-Host -ForegroundColor Green "Installing Company Portal"
$packagePath    = "C:\OSDCloud\CompanyPortal\CompanyPortal.appxbundle"
$dependencyPath = "C:\OSDCloud\CompanyPortal\Dependencies"
$dependencies   = Get-ChildItem -Path $dependencyPath -Filter *.appx | ForEach-Object { $_.FullName }

Add-AppxProvisionedPackage -Online -PackagePath $packagePath -DependencyPackagePath $dependencies -SkipLicense

