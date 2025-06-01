# --------------------------------------------------------------------------------------------------------
# Parameters
# --------------------------------------------------------------------------------------------------------
Param(
    [Parameter(Mandatory)]
    [string]$IntuneWinFile,

    [Parameter(Mandatory)]
    [string]$DisplayName,

    [Parameter(Mandatory)]
    [string]$AppType
)

# --------------------------------------------------------------------------------------------------------
# Tenant Configuration
# --------------------------------------------------------------------------------------------------------

$myTenantID = "tugicloud.onmicrosoft.com"
$myClientID = "9a23bd51-2de6-46fd-ba90-c28bdca28a20"
$myClientSecret = "db617388-62c4-4fea-86dd-241865e4f741"

# --------------------------------------------------------------------------------------------------------
# Main Script
# --------------------------------------------------------------------------------------------------------

# Set Execution Policy
Set-ExecutionPolicy ByPass -Scope CurrentUser

# Check that the app registration is defined correctly in the script
If ($myTenantID -eq "" -or $myClientID -eq "" -or $myClientSecret -eq "")
{
    CLS
    Write-Host "Please configure your tenant information (Azure > App Registration) in the script first:"
    Write-Host "`n$($MyInvocation.MyCommand.Path)"
    Write-Host "`nCheck out GitHub for more details: https://github.com/tugich/SetupPackagerIntune"
    Exit 1
}

# Connect to MS Intune
If (Get-InstalledModule "IntuneWin32App")
{
  Connect-MSIntuneGraph -TenantID $myTenantID -ClientID $myClientID -ClientSecret $myClientSecret
}
Else
{
  CLS
  Write-Host "IntuneWin32App module not found - please install it first." -ForegroundColor Black -BackgroundColor Yellow
  Write-Host "@Link: https://github.com/MSEndpointMgr/IntuneWin32App"
  # Install-Module -Name "IntuneWin32App" -RequiredVersion 1.4.1
  Exit 1
}

# Get MSI meta data from .intunewin file
$IntuneWinMetaData = Get-IntuneWin32AppMetaData -FilePath $IntuneWinFile

If ($IntuneWinMetaData.ApplicationInfo.MsiInfo)
{
    # Create requirement rule for all platforms and Windows 11
    $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease "W11_22H2"

    # Create PowerShell script detection rule
    $DetectionRule = New-IntuneWin32AppDetectionRuleMSI -ProductCode $IntuneWinMetaData.ApplicationInfo.MsiInfo.MsiProductCode -ProductVersionOperator "greaterThanOrEqual" -ProductVersion $IntuneWinMetaData.ApplicationInfo.MsiInfo.MsiProductVersion

    # Create custom return code
    $ReturnCode = New-IntuneWin32AppReturnCode -ReturnCode 1337 -Type "retry"

    # Add new EXE Win32 app
    $InstallCommandLine = 'msiexec /i "' + $($IntuneWinMetaData.ApplicationInfo.SetupFile) + '" /qn'
    $UninstallCommandLine = 'msiexec /x "' + $($IntuneWinMetaData.ApplicationInfo.MsiInfo.MsiProductCode) + '" /qn'

    Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description "Imported with Setup Packager for Intune - by TUGI" -Publisher $IntuneWinMetaData.ApplicationInfo.MsiInfo.MsiPublisher -InstallExperience "system" -AppVersion $IntuneWinMetaData.ApplicationInfo.MsiInfo.MsiProductVersion -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -ReturnCode $ReturnCode -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Verbose
    Exit 0
}
Else
{
    CLS
    Write-Host "Upload failed: Sorry, MSI only supported."
    Exit 1
}
