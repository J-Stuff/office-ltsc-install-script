#Requires -RunAsAdministrator
# Office 2024 LTSC Uninstaller script
# Made by J Stuff // jstuff@j-stuff.net // github.com/JStuff

Clear-Host
# ==============================================
# Script Configs

$installerPath = ""
$removeConfigPath = ""
# ==============================================

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
    Write-Host "Removing..." -ForegroundColor Yellow
    Write-Host "(This will take a while, see your IT administrator if this does not complete within 20 minutes.)" -ForegroundColor DarkGray
    Write-Host "Removal started at: $(Get-Date -DisplayHint Time)" -ForegroundColor DarkGray
    Start-Process -FilePath $installerPath -ArgumentList "/configure `"$removeConfigPath`"" -Wait -NoNewWindow
}

# Execution Logic
Write-Host "Office 2024 - LTSC Uninstall Script"
KillAnyRunningOfficeApps
Uninstall-ExistingOffice
