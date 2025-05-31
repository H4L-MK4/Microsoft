<#
.SYNOPSIS
Office 365 PowerShell Modules Installer

.DESCRIPTION
This script installs and updates the Office 365 PowerShell modules. You can also connect to the services.

.AUTHOR
H4L-MK4

.VERSION
1.1 - Removed comments and simplified the script

#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script requires Administrator privileges. Please re-run PowerShell as Administrator."
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

$currentPolicy = Get-ExecutionPolicy
if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
    Write-Host "Temporarily setting ExecutionPolicy to RemoteSigned for this session..." -ForegroundColor Yellow
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction Stop
    } catch {
        Write-Error "Failed to set ExecutionPolicy. Please manually run: Set-ExecutionPolicy RemoteSigned -Scope Process -Force"
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

$Global:AttemptedConnections = @{}

$modules = @(
    @{ Name = "Microsoft.Graph"; Description = "Microsoft Graph SDK"; ConnectCmdletString = "Connect-MgGraph" },
    @{ Name = "ExchangeOnlineManagement"; Description = "Exchange Online Management"; ConnectCmdletString = "Connect-ExchangeOnline" },
    @{ Name = "Microsoft.Online.SharePoint.PowerShell"; Description = "SharePoint Online Management"; ConnectCmdletString = "Connect-SPOService" },
    @{ Name = "MicrosoftTeams"; Description = "Microsoft Teams"; ConnectCmdletString = "Connect-MicrosoftTeams" },
    @{ Name = "AzureAD"; Description = "Azure Active Directory (Legacy)"; ConnectCmdletString = "Connect-AzureAD" },
    @{ Name = "MSOnline"; Description = "MSOnline (Legacy for MFA/Licensing)"; ConnectCmdletString = "Connect-MsolService" },
    @{ Name = "Microsoft.PowerApps.Administration.PowerShell"; Description = "Power Platform Administration"; ConnectCmdletString = "Add-PowerAppsAccount -Endpoint admin" },
    @{ Name = "Microsoft.PowerApps.PowerShell"; Description = "Power Apps Maker"; ConnectCmdletString = "Add-PowerAppsAccount" },
    @{ Name = "Az.Accounts"; Description = "Azure Core Accounts"; ConnectCmdletString = "Connect-AzAccount" },
    @{ Name = "PnP.PowerShell"; Description = "PnP PowerShell (SharePoint/Teams)"; ConnectCmdletString = "Connect-PnPOnline" }
)
function Show-Menu {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "   Office 365 PowerShell Module Installer " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host

    for ($i = 0; $i -lt $modules.Count; $i++) {
        $moduleItem = $modules[$i]
        $foundModules = Get-Module -ListAvailable -Name $moduleItem.Name -ErrorAction SilentlyContinue
        
        if ($foundModules) {
            $latestVersion = ($foundModules | Sort-Object -Property Version -Descending | Select-Object -First 1).Version
            $statusText = "[INSTALLED - v$latestVersion]"
            $lineColor = "Green"
        } else {
            $statusText = "[NOT INSTALLED]"
            $lineColor = "White"
        }
        Write-Host ("{0,2}. {1,-45} {2}" -f ($i + 1), $moduleItem.Description, $statusText) -ForegroundColor $lineColor
    }

    Write-Host
    Write-Host "[A] Install ALL Not Installed Modules" -ForegroundColor Magenta
    Write-Host "[U] Update ALL Installed Modules" -ForegroundColor Blue
    Write-Host "[C] Connect to Services" -ForegroundColor Yellow
    Write-Host "[S] Show Installation Status (Refresh)" -ForegroundColor Cyan
    Write-Host "[Q] Quit" -ForegroundColor Red
    Write-Host
}
function Manage-Module {
    param (
        [string]$ModuleName,
        [string]$Action = "Install"
    )

    $actionVerb = if ($Action -eq "Install") { "Installing" } else { "Updating" }
    Write-Host "$actionVerb module '$ModuleName'..." -ForegroundColor Yellow
    
    try {
        if ($Action -eq "Install") {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        } elseif ($Action -eq "Update") {
            Update-Module -Name $ModuleName -Force -Scope CurrentUser -ErrorAction Stop
        }
        Write-Host "Successfully $($actionVerb.ToLower() -replace 'ing$', 'ed') '$ModuleName'." -ForegroundColor Green
    } catch {
        Write-Error "Failed to $($Action.ToLower()) module '$ModuleName'. Error: $($_.Exception.Message)"
    }
    Start-Sleep -Milliseconds 250
}
function Show-ConnectServicesMenu {
    param (
        [parameter(Mandatory=$true)]
        [System.Collections.IList]$ConnectableModules 
    )
    Clear-Host
    Write-Host "==================================================" -ForegroundColor DarkCyan
    Write-Host "           Connect to Office 365 Services" -ForegroundColor DarkCyan
    Write-Host "==================================================" -ForegroundColor DarkCyan
    Write-Host

    $menuIndex = 1
    foreach ($moduleItem in $ConnectableModules) {
        $status = "[Not Attempted]"
        $color = "White"
        if ($Global:AttemptedConnections.ContainsKey($moduleItem.Name)) {
            if ($Global:AttemptedConnections[$moduleItem.Name].Success) {
                $status = "[CONNECTED]"
                $color = "Green"
            } else {
                $status = "[ATTEMPT FAILED]"
                $color = "Red"
            }
        }
        Write-Host ("{0,2}. Connect to {1,-35} {2}" -f $menuIndex, $moduleItem.Description, $status) -ForegroundColor $color
        $menuIndex++
    }
    Write-Host
    Write-Host "[B] Back to Main Menu" -ForegroundColor Yellow
    Write-Host
}
function Invoke-ModuleConnect {
    param (
        [parameter(Mandatory=$true)]
        [hashtable]$ModuleToConnect
    )

    Write-Host ("Attempting to connect to '{0}'..." -f $ModuleToConnect.Description) -ForegroundColor Yellow
    $connectCommand = $ModuleToConnect.ConnectCmdletString
    $success = $false

    try {
        if ($ModuleToConnect.Name -eq "Microsoft.Online.SharePoint.PowerShell") {
            $spAdminUrl = Read-Host -Prompt "Enter SharePoint Admin Center URL (e.g., https://yourtenant-admin.sharepoint.com)"
            if (-not [string]::IsNullOrWhiteSpace($spAdminUrl)) {
                $connectCommand = "Connect-SPOService -Url '$spAdminUrl'"
                Invoke-Expression $connectCommand
                $success = $true
            } else {
                Write-Warning "SharePoint Admin URL not provided. Skipping connection."
            }
        } elseif ($ModuleToConnect.Name -eq "PnP.PowerShell") {
            $pnpSiteUrl = Read-Host -Prompt "Enter SharePoint Site URL for PnP Connection (e.g., https://yourtenant.sharepoint.com/sites/your-site)"
            if (-not [string]::IsNullOrWhiteSpace($pnpSiteUrl)) {
                $connectCommand = "Connect-PnPOnline -Url '$pnpSiteUrl' -Interactive"
                Invoke-Expression $connectCommand
                $success = $true
            } else {
                Write-Warning "PnP Site URL not provided. Skipping connection."
            }
        } else {
            Invoke-Expression $connectCommand
            $success = $true 
        }

        if ($success) {
            Write-Host ("Successfully initiated connection to '{0}'." -f $ModuleToConnect.Description) -ForegroundColor Green
            $Global:AttemptedConnections[$ModuleToConnect.Name] = @{ Success = $true; Timestamp = Get-Date }
        }
    } catch {
        Write-Error ("Failed to connect to '{0}'. Error: $($_.Exception.Message)" -f $ModuleToConnect.Description)
        $Global:AttemptedConnections[$ModuleToConnect.Name] = @{ Success = $false; Timestamp = Get-Date; ErrorMessage = $_.Exception.Message }
    }
    Start-Sleep -Milliseconds 100 
}

do {
    Show-Menu
    $userInput = Read-Host -Prompt "Enter your choice"

    switch ($userInput.ToUpper()) {
        "A" {
            Write-Host "Installing all not installed modules..." -ForegroundColor Magenta
            foreach ($moduleEntry in $modules) {
                if (-not (Get-Module -ListAvailable -Name $moduleEntry.Name -ErrorAction SilentlyContinue)) {
                    Manage-Module -ModuleName $moduleEntry.Name -Action "Install"
                }
            }
            Write-Host "Installation process completed." -ForegroundColor Magenta
        }
        "U" {
            Write-Host "Updating all installed modules..." -ForegroundColor Blue
            foreach ($moduleEntry in $modules) {
                if (Get-Module -ListAvailable -Name $moduleEntry.Name -ErrorAction SilentlyContinue) {
                    Manage-Module -ModuleName $moduleEntry.Name -Action "Update"
                }
            }
            Write-Host "Update process completed." -ForegroundColor Blue
        }
        "C" {
            [array]$connectableModulesList = $modules | Where-Object { 
                ($_.ConnectCmdletString -ne $null) -and
                ($_.ConnectCmdletString -is [string]) -and
                ($_.ConnectCmdletString.Trim().Length -gt 0)
            }
            if (-not $connectableModulesList) { 
                Write-Warning "No modules are currently configured with a valid Connect command. Please check the script's module definitions."
                Write-Host "Press any key to return to the main menu..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            } else {
                                do {
                    Show-ConnectServicesMenu -ConnectableModules $connectableModulesList
                    $connectInput = Read-Host -Prompt "Select service to connect to, or [B] for back"
                    if ($connectInput.ToUpper() -eq 'B') { break }

                    if ($connectInput -match "^\d+$") {
                        $connectSelectionIndex = [int]$connectInput - 1
                        if ($connectSelectionIndex -ge 0 -and $connectSelectionIndex -lt $connectableModulesList.Count) {
                            Invoke-ModuleConnect -ModuleToConnect $connectableModulesList[$connectSelectionIndex]
                            Write-Host "Press any key to return to the Connect menu..."
                            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                        } else {
                            Write-Warning "Invalid selection number. Please choose from the listed options."
                            Start-Sleep -Seconds 2
                        }
                    } else {
                        Write-Warning "Invalid input. Please enter a number or 'B'."
                        Start-Sleep -Seconds 2
                    }
                } while ($true)
            }
        }
        "S" { 
            Write-Host "Refreshing status..." -ForegroundColor Cyan 
        }
        "Q" {
            Write-Host "Exiting script." -ForegroundColor Green
            exit
        }
        default {
            if ($userInput -match "^\d+$") {
                $selectionIndex = [int]$userInput - 1
                if ($selectionIndex -ge 0 -and $selectionIndex -lt $modules.Count) {
                    $selectedModule = $modules[$selectionIndex]
                    $installed = Get-Module -ListAvailable -Name $selectedModule.Name -ErrorAction SilentlyContinue
                    if ($installed) {
                        Manage-Module -ModuleName $selectedModule.Name -Action "Update"
                    } else {
                        Manage-Module -ModuleName $selectedModule.Name -Action "Install"
                    }
                } else {
                    Write-Warning "Invalid selection number."
                }
            } else {
                Write-Warning "Invalid input. Please enter a number or a menu letter."
            }
        }
    }

    if ($userInput.ToUpper() -ne "Q" -and $userInput.ToUpper() -ne "C") { 
        Write-Host "Press any key to return to the menu..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }

} while ($true) 