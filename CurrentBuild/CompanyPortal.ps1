# InstallCompanyPortal.ps1
Write-Host -ForegroundColor Green "Provisioning Company Portal for all users…"

$bundle       = "C:\OSDCloud\CompanyPortal\CompanyPortal.appxbundle"
$depFolder    = "C:\OSDCloud\CompanyPortal\Dependencies"
$dependencies = Get-ChildItem -Path $depFolder -Filter *.appx | ForEach-Object FullName

# First add dependency frameworks
foreach($dep in $dependencies) {
  Add-AppxProvisionedPackage -Online `
    -PackagePath $dep `
    -SkipLicense
}

# Then the main bundle
Add-AppxProvisionedPackage -Online `
  -PackagePath $bundle `
  -DependencyPackagePath $dependencies `
  -SkipLicense
