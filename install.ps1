#Requires -RunAsAdministrator
# Office 2024 LTSC Installer script
# Made by github.com/J-Stuff // jstuff@j-stuff.net
# Licenced under the GNU General Public License v3.0

# Keep me here so any syntax errors aren't cleared below
Clear-Host
# ==============================================
# Script Configs
# Path to folder containing your Office ODT setup.exe, & installer configs
$installerFolder = ""
# Path to Office ODT setup.exe
$installerPath = ""
# Path to .xml config for Office ODT (generate one at https://config.office.com/officeSettings/configurations)
$officeConfigPath = ""
# Path to .xml config to remove all office applications (use this config: https://learn.microsoft.com/en-us/answers/questions/1165399/uninstall-office-365-with-odt#:~:text=If%20you%20want%20to%20remove%20all%20C2R%20versions%20of%20Office%20products%2C%20please%20try%20following%20.xml%20file.)
$removeConfigPath = ""
# ==============================================

# (!) Don't touch past this point unless you know what you are doing (!)

$officeInstallPath = "C:\Program Files\Microsoft Office\Office16"

function Test-Office2024Installed {
    # Checks for both x64 and old x86 installations
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $apps = Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Office*2024*" }
    
    return [bool]$apps
}

function Test-Paths {
    $pathsToTest = $installerFolder, $installerPath, $officeConfigPath, $removeConfigPath
    $allPathsExists = (Test-Path -Path $pathsToTest) -notcontains $false
    if (-not ($allPathsExists)) {
        Write-Host "(!) One or more installer or config paths are invalid in this script's config. Please see your Admininstrator (!)" -ForegroundColor DarkRed
        Start-Sleep 10
        exit 1
    }
}

function KillAnyRunningOfficeApps {
    Write-Host "Closing any running Office Application instances" -ForegroundColor Yellow
    $apps = @('winword','excel','outlook','powerpnt')
    $apps | ForEach-Object {
        $gp = Get-Process -Name "$_*" -ErrorAction SilentlyContinue
        if ($gp) {
            Stop-Process -Name "$_*" -Force
            Write-Host "$_ was closed" -ForegroundColor DarkRed
        }
    }
}

function Uninstall-ExistingOffice {
    # This will clean ALL versions of office found on the device!
    Write-Host "Removing existing Microsoft Office Installations on this PC" -ForegroundColor Yellow
    Write-Host "(This could take us a moment)" -ForegroundColor DarkGray
    Start-Process -FilePath $installerPath -ArgumentList "/configure `"$removeConfigPath`"" -Wait -NoNewWindow
}

function Install-Office2024 {
    Write-Host "Installing Office 2024 - LTSC..."
    Write-Host "(This will take a while, see your IT administrator if this does not complete within 20 minutes.)" -ForegroundColor DarkGray
    Write-Host "Installation started at: $(Get-Date -DisplayHint Time)" -ForegroundColor DarkGray
    Set-Location $installerFolder
    Start-Process -FilePath $installerPath -ArgumentList "/configure `"$officeConfigPath`"" -Wait -NoNewWindow
    Write-Host "Finalizing office installation..."
    # Sleep to allow Office Installer to actually finish, as the installer lies a bit and returns complete while the last few steps are still ongoing. 
    # This causes any future steps to sometimes fail if Word, etc are not in a Ready state yet. Which often happens on slow/overloaded devices.
    Start-Sleep 20
    Write-Host "Office is installed" -ForegroundColor Green

}

function fetchActivationStatus {
    Write-Host "Office Activation Status:"
    Set-Location $officeInstallPath
    $result = cscript //nologo ospp.vbs /dstatus
    if ($result -like "*LICENSE STATUS:  ---LICENSED---*") {
        Write-Host "Office 2024 - LTSC is activated" -ForegroundColor Green
    } else {
        Write-Host "Office 2024 - LTSC is NOT activated!" -ForegroundColor DarkRed
        Write-Host "Please see your IT Administrator!" -ForegroundColor DarkRed
        Start-Sleep 5
    }
}

# Execution Logic
Write-Host "Office 2024 - LTSC Install script"
Test-Paths

if (-not (Test-Office2024Installed)) {
    KillAnyRunningOfficeApps
    Uninstall-ExistingOffice
    Install-Office2024
    fetchActivationStatus
    Start-Sleep 3
    exit 0
} else {
    Write-Host "Office 2024 - LTSC is already installed :)" -ForegroundColor Green
    fetchActivationStatus
    Start-Sleep 3
    exit 0
}