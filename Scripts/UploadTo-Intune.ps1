# Parameters
Param(
    [Parameter(Mandatory)]
    [string]$Tenant,

    [Parameter(Mandatory)]
    [string]$IntuneWinFile,

    [Parameter(Mandatory)]
    [string]$DisplayName
)

# Set Execution Policy
Set-ExecutionPolicy ByPass -Scope CurrentUser

# Connect to Intune
If (Get-InstalledModule "IntuneWin32App")
{
  Connect-MSIntuneGraph -TenantID $Tenant | Out-Null
}
Else
{
  Write-Host "IntuneWin32App module not found - please install it first." -ForegroundColor Black -BackgroundColor Yellow
  Write-Host "https://github.com/MSEndpointMgr/IntuneWin32App"
  # Install-Module -Name "IntuneWin32App"
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

    Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description "Imported with Setup Packager for Intune - by TUGI" -Publisher "SetupPackager" -InstallExperience "system" -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -ReturnCode $ReturnCode -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Verbose
    Exit 0
}
Else
{
    CLS
    Write-Host "Upload failed: Sorry, MSI only supported."
    Exit 1
}
