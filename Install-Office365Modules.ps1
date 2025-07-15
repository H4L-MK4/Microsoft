<#
.SYNOPSIS
Office 365 PowerShell Modules Installer

.DESCRIPTION
Installs/updates Office 365 PowerShell modules, shows status, and opens an interactive browser sign-in for each service.

.AUTHOR
H4L-MK4

.VERSION
1.2 – updated design and layout

#>

$esc = [char]27
$KULLColors = @{
    Reset         = "$esc[0m"
    Border        = "$esc[38;2;160;160;160m" # KULL's gray: #A0A0A0
    Title         = "$esc[38;2;255;199;153m" # KULL's bright yellow: #FFC799
    Header        = "$esc[38;2;230;185;157m" # KULL's yellow: #E6B99D
    Success       = "$esc[38;2;53;223;177m"  # KULL's bright green: #35DFB1
    Failure       = "$esc[38;2;255;128;128m" # KULL's bright red: #FF8080
    Warning       = "$esc[38;2;255;199;153m" # KULL's bright yellow: #FFC799
    Text          = "$esc[38;2;255;255;255m" # KULL's foreground: #FFFFFF
    Muted         = "$esc[38;2;160;160;160m" # KULL's gray: #A0A0A0
    Connected     = "$esc[38;2;53;223;177m"  # Green for connected status
    FailedConnect = "$esc[38;2;255;128;128m" # Red for failed status
    MenuKey       = "$esc[38;2;255;199;153m" # KULL's bright yellow: #FFC799
    MenuText      = "$esc[38;2;153;255;228m" # KULL's cyan: #99FFE4
    Version       = "$esc[38;2;255;199;153m" # KULL's bright yellow: #FFC799
}

function Write-Centered {
    param(
        [string]$Text,
        [int]$Width = [Console]::WindowWidth
    )
    $visibleText = $Text -replace "$esc\[[0-9;]*m"
    $paddingLength = [math]::Max(0, [math]::Floor(($Width - $visibleText.Length) / 2))
    $padding = ' ' * $paddingLength
    
    Write-Host "$padding$Text"
}

function Show-Header {
    $width = [Console]::WindowWidth
    $border = '═' * $width
    Write-Host "$($KULLColors.Border)$border$($KULLColors.Reset)"
    Write-Centered "$($KULLColors.Title)Office 365 PowerShell Module Installer$($KULLColors.Reset)"
    Write-Host "$($KULLColors.Border)$border$($KULLColors.Reset)"
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
     ).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Write-Warning 'This script requires Administrator privileges. Please re-run PowerShell as Administrator.'
    exit 1
}

$modules = @(
    @{ Name = 'Microsoft.Graph';                          Description = 'Microsoft Graph SDK';          ConnectCmdletString = 'Connect-MgGraph' },
    @{ Name = 'ExchangeOnlineManagement';                  Description = 'Exchange Online Management';  ConnectCmdletString = 'Connect-ExchangeOnline' },
    @{ Name = 'Microsoft.Online.SharePoint.PowerShell';     Description = 'SharePoint Online Mgmt';      ConnectCmdletString = 'Connect-SPOService' },
    @{ Name = 'MicrosoftTeams';                            Description = 'Microsoft Teams';            ConnectCmdletString = 'Connect-MicrosoftTeams' },
    @{ Name = 'AzureAD';                                    Description = 'Azure AD (Legacy)';          ConnectCmdletString = 'Connect-AzureAD' },
    @{ Name = 'MSOnline';                                  Description = 'MSOnline (Legacy)';          ConnectCmdletString = 'Connect-MsolService' },
    @{ Name = 'Microsoft.PowerApps.Administration.PowerShell'; Description = 'Power Apps Admin';       ConnectCmdletString = 'Add-PowerAppsAccount -Endpoint admin' },
    @{ Name = 'Microsoft.PowerApps.PowerShell';            Description = 'Power Apps Maker';           ConnectCmdletString = 'Add-PowerAppsAccount' },
    @{ Name = 'Az.Accounts';                               Description = 'Azure Core Accounts';        ConnectCmdletString = 'Connect-AzAccount' },
    @{ Name = 'PnP.PowerShell';                            Description = 'PnP PowerShell';             ConnectCmdletString = 'Connect-PnPOnline' }
)

function Show-ModuleStatus {
    $headerColor = $KULLColors.Header
    $versionColor = $KULLColors.Version
    $reset = $KULLColors.Reset

    $statusData = $modules | ForEach-Object {
        $name = $_.Name
        $mod  = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue |
                Sort-Object Version -Descending | Select-Object -First 1
        
        [PSCustomObject]@{
            Module    = $name
            Installed = if ($mod) { 'Yes' } else { 'No' }
            Version   = if ($mod) { $mod.Version.ToString() } else { '-' }
        }
    }

    $maxModuleLen = ($statusData.Module | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $maxInstalledLen = ($statusData.Installed | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $maxModuleLen = [math]::Max($maxModuleLen, 'Module'.Length)
    $maxInstalledLen = [math]::Max($maxInstalledLen, 'Installed'.Length)

    $header = "{0,-$maxModuleLen}  {1,-$maxInstalledLen}  {2}" -f 'Module', 'Installed', 'Version'
    Write-Host "`n$headerColor$header$reset"
    $separator = "{0}  {1}  {2}" -f ('-' * $maxModuleLen), ('-' * $maxInstalledLen), ('-' * 'Version'.Length)
    Write-Host "$headerColor$separator$reset"

    foreach ($item in $statusData) {
        $versionText = if ($item.Version -ne '-') { "$versionColor$($item.Version)$reset" } else { '-' }
        $line = "{0,-$maxModuleLen}  {1,-$maxInstalledLen}  {2}" -f $item.Module, $item.Installed, $versionText
        Write-Host $line
    }
    Write-Host ''
}

function Get-LatestModuleVersion {
    param([string]$ModuleName)
    try {
        $latest = Find-Module -Name $ModuleName -ErrorAction Stop | Select-Object -First 1
        return $latest.Version
    } catch {
        return $null
    }
}

function Manage-Module {
    param (
        [Parameter(Mandatory)] [string] $ModuleName,
        [Parameter(Mandatory)][ValidateSet('Install', 'Update', 'Uninstall')] [string] $Action
    )
    $actionVerbs = @{
        Install   = @{ Present = 'Installing'; Past = 'installed' }
        Update    = @{ Present = 'Updating';   Past = 'updated' }
        Uninstall = @{ Present = 'Uninstalling'; Past = 'uninstalled' }
    }
    $presentVerb = $actionVerbs[$Action].Present
    $pastVerb = $actionVerbs[$Action].Past

    Write-Centered "$($KULLColors.Warning)$presentVerb module '$ModuleName'...$($KULLColors.Reset)"
    try {
        switch ($Action) {
            'Install'   { Install-Module -Name $ModuleName -Force -AllowClobber -Scope AllUsers -ErrorAction Stop }
            'Update'    { Update-Module -Name $ModuleName -Force -Scope AllUsers -ErrorAction Stop }
            'Uninstall' { Uninstall-Module -Name $ModuleName -Force -ErrorAction Stop }
        }
        Write-Centered "$($KULLColors.Success)Successfully $pastVerb '$ModuleName'.$($KULLColors.Reset)"
    } catch {
        $errorMessage = "Failed to $($Action.ToLower()) '$ModuleName': $($_.Exception.Message)"
        Write-Centered "$($KULLColors.Failure)$errorMessage$($KULLColors.Reset)"
    }
}

function Show-Menu {
    Clear-Host
    Show-Header
    Write-Host
    $keyColor = $KULLColors.MenuKey
    $textColor = $KULLColors.MenuText
    $reset = $KULLColors.Reset

    Write-Centered "${keyColor}S)$reset $textColor Show module status$reset"
    Write-Centered "${keyColor}I)$reset $textColor Install missing module(s)$reset"
    Write-Centered "${keyColor}U)$reset $textColor Update installed module(s)$reset"
    Write-Centered "${keyColor}X)$reset $textColor Uninstall installed module(s)$reset"
    Write-Centered "${keyColor}C)$reset $textColor Connect to a service$reset"
    Write-Centered "${keyColor}Q)$reset $textColor Quit$reset"
    Write-Host
}

function Invoke-ModuleConnect {
    param ([hashtable] $ModuleToConnect)

    Write-Centered "$($KULLColors.Warning)Connecting to '$($ModuleToConnect.Description)'...$($KULLColors.Reset)"
    $success = $false

    try {
        switch ($ModuleToConnect.Name) {
            'Microsoft.Online.SharePoint.PowerShell' {
                $url = Read-Host 'Enter SharePoint Admin Center URL'
                if ($url) {
                    Connect-SPOService -Url $url -ErrorAction Stop
                    $success = $true
                }
            }
            'PnP.PowerShell' {
                $url = Read-Host 'Enter site URL for PnP connection'
                if ($url) {
                    Connect-PnPOnline -Url $url -Interactive -ErrorAction Stop
                    $success = $true
                }
            }
            'ExchangeOnlineManagement' {
                Write-Host 'Opening browser for interactive sign-in…' -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
                $success = $true
            }
            default {
                Invoke-Expression $ModuleToConnect.ConnectCmdletString
                $success = $true
            }
        }
        if ($success) {
            Write-Centered "$($KULLColors.Success)Connected to '$($ModuleToConnect.Description)'.$($KULLColors.Reset)"
            $Global:AttemptedConnections[$ModuleToConnect.Name] = @{ Success = $true; Time = Get-Date }
        } else {
            Write-Centered "$($KULLColors.Muted)Connection skipped (no input provided).$($KULLColors.Reset)"
        }
    } catch {
        $errorMessage = "Failed to connect to '$($ModuleToConnect.Description)': $($_.Exception.Message)"
        Write-Centered "$($KULLColors.Failure)$errorMessage$($KULLColors.Reset)"
        $Global:AttemptedConnections[$ModuleToConnect.Name] = @{ Success = $false; Time = Get-Date; Error = $_.Exception.Message }
    }
    Start-Sleep -Milliseconds 100
}

# Main loop
$Global:AttemptedConnections = @{}
$exitLoop = $false

do {
    Show-Menu
    $choice = Read-Host "$($KULLColors.Muted)Enter choice:$($KULLColors.Reset)"

    switch ($choice.ToUpper()) {
        'S' {
            Clear-Host
            Show-Header
            Show-ModuleStatus
            Write-Centered "$($KULLColors.Muted)Press any key to return...$($KULLColors.Reset)"
            [void][System.Console]::ReadKey($true)
        }
        'I' {
            [array]$missingModules = $modules | Where-Object { -not (Get-Module -ListAvailable -Name $_.Name -ErrorAction SilentlyContinue) }
            if ($missingModules.Count -eq 0) {
                Clear-Host
                Show-Header
                Write-Centered "$($KULLColors.Success)All modules are already installed.$($KULLColors.Reset)"; Start-Sleep 2
                continue
            }
            
            do {
                Clear-Host
                Show-Header
                Write-Host
                Write-Centered "$($KULLColors.Header)--- Install Modules ---$($KULLColors.Reset)"
                Write-Host
                $keyColor = $KULLColors.MenuKey
                $textColor = $KULLColors.MenuText
                $reset = $KULLColors.Reset
                for ($i = 0; $i -lt $missingModules.Count; $i++) {
                    Write-Centered "${keyColor}$($i + 1))$reset $textColor Install $($missingModules[$i].Description)$reset"
                }
                Write-Host
                Write-Centered "${keyColor}A)$reset $textColor Install ALL missing modules$reset"
                Write-Host

                $sel = Read-Host "$($KULLColors.Muted)Select an option, or press [Enter] to go back$($KULLColors.Reset)"

                if ($sel -eq '') { break }
                
                if ($sel.ToUpper() -eq 'A') {
                    foreach ($m in $missingModules) {
                        Manage-Module -ModuleName $m.Name -Action Install
                    }
                    Write-Centered "$($KULLColors.Success)All missing modules installed. Press any key to continue...$($KULLColors.Reset)"
                    [void][System.Console]::ReadKey($true)
                    break # Exit to main menu
                } elseif ($sel -match '^\d+$' -and ($sel -as [int]) -ge 1 -and ($sel -as [int]) -le $missingModules.Count) {
                    $moduleToInstall = $missingModules[($sel -as [int]) - 1]
                    Manage-Module -ModuleName $moduleToInstall.Name -Action Install
                    Write-Centered "$($KULLColors.Muted)Press any key to continue...$($KULLColors.Reset)"
                    [void][System.Console]::ReadKey($true)
                    # Refresh list in case we want to install another
                    $missingModules = $modules | Where-Object { -not (Get-Module -ListAvailable -Name $_.Name -ErrorAction SilentlyContinue) }
                    if ($missingModules.Count -eq 0) {
                         Write-Centered "$($KULLColors.Success)All modules are now installed. Press any key to continue...$($KULLColors.Reset)"
                         [void][System.Console]::ReadKey($true)
                         break # Exit to main menu
                    }
                } else {
                    Write-Centered "$($KULLColors.Failure)Invalid selection$($KULLColors.Reset)"; Start-Sleep 1
                }
            } while ($true)
        }
        'U' {
            [array]$installedModules = $modules | Where-Object { Get-Module -ListAvailable -Name $_.Name -ErrorAction SilentlyContinue }
            if ($installedModules.Count -eq 0) {
                Clear-Host
                Show-Header
                Write-Centered "$($KULLColors.Warning)No modules are installed to update.$($KULLColors.Reset)"; Start-Sleep 2
                continue
            }
            
            do {
                Clear-Host
                Show-Header
                Write-Host
                $moduleDetails = foreach ($m in $installedModules) {
                    $mod = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
                    $latest = Get-LatestModuleVersion -ModuleName $m.Name
                    [pscustomobject]@{
                        Name          = $m.Name
                        Description   = $m.Description
                        Version       = $mod.Version
                        LatestVersion = $latest
                    }
                }

                Write-Centered "$($KULLColors.Header)--- Update Modules ---$($KULLColors.Reset)"
                Write-Host
                $keyColor = $KULLColors.MenuKey
                $textColor = $KULLColors.MenuText
                $versionColor = $KULLColors.Version
                $reset = $KULLColors.Reset
                for ($i = 0; $i -lt $moduleDetails.Count; $i++) {
                    $detail = $moduleDetails[$i]
                    $versionText = "$textColor(installed v$versionColor$($detail.Version)$textColor,$reset $textColorlatest v$versionColor$($detail.LatestVersion)$textColor)$reset"
                    Write-Centered "${keyColor}$($i + 1))$reset $textColor Update $($detail.Description) $versionText"
                }
                Write-Host
                Write-Centered "${keyColor}A)$reset $textColor Update ALL installed modules$reset"
                Write-Host

                $sel = Read-Host "$($KULLColors.Muted)Select an option, or press [Enter] to go back$($KULLColors.Reset)"
                
                if ($sel -eq '') { break }

                if ($sel.ToUpper() -eq 'A') {
                    foreach ($m in $installedModules) {
                        Manage-Module -ModuleName $m.Name -Action Update
                    }
                    Write-Centered "$($KULLColors.Success)All installed modules updated. Press any key to continue...$($KULLColors.Reset)"
                    [void][System.Console]::ReadKey($true)
                    break
                } elseif ($sel -match '^\d+$' -and ($sel -as [int]) -ge 1 -and ($sel -as [int]) -le $installedModules.Count) {
                    $moduleToUpdate = $installedModules[($sel -as [int]) - 1]
                    Manage-Module -ModuleName $moduleToUpdate.Name -Action Update
                    Write-Centered "$($KULLColors.Muted)Press any key to continue...$($KULLColors.Reset)"
                    [void][System.Console]::ReadKey($true)
                } else {
                    Write-Centered "$($KULLColors.Failure)Invalid selection$($KULLColors.Reset)"; Start-Sleep 1
                }
            } while ($true)
        }
        'X' {
            [array]$installedModules = $modules | Where-Object { Get-Module -ListAvailable -Name $_.Name -ErrorAction SilentlyContinue }
            if ($installedModules.Count -eq 0) {
                Clear-Host
                Show-Header
                Write-Centered "$($KULLColors.Warning)No modules are installed.$($KULLColors.Reset)"; Start-Sleep 2
                continue
            }
            
            do {
                Clear-Host
                Show-Header
                Write-Host
                $moduleDetails = foreach ($m in $installedModules) {
                    $mod = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
                    [pscustomobject]@{
                        Name        = $m.Name
                        Description = $m.Description
                        Version     = $mod.Version
                    }
                }

                Write-Centered "$($KULLColors.Header)--- Uninstall Modules ---$($KULLColors.Reset)"
                Write-Host
                $keyColor = $KULLColors.MenuKey
                $textColor = $KULLColors.MenuText
                $versionColor = $KULLColors.Version
                $reset = $KULLColors.Reset
                for ($i = 0; $i -lt $moduleDetails.Count; $i++) {
                    $detail = $moduleDetails[$i]
                    $versionText = "$textColor(v$versionColor$($detail.Version)$textColor)$reset"
                    Write-Centered "${keyColor}$($i + 1))$reset $textColor Uninstall $($detail.Description) $versionText"
                }
                Write-Host
                Write-Centered "${keyColor}A)$reset $textColor Uninstall ALL installed modules$reset"
                Write-Host

                $sel = Read-Host "$($KULLColors.Muted)Select an option, or press [Enter] to go back$($KULLColors.Reset)"
                
                if ($sel -eq '') { break }

                if ($sel.ToUpper() -eq 'A') {
                    $confirmPrompt = "$($KULLColors.Warning)ARE YOU SURE you want to uninstall ALL modules? This cannot be undone. (y/n)$($KULLColors.Reset)"
                    $userInput = Read-Host $confirmPrompt
                    if ($userInput -ne 'y') {
                        Write-Centered "$($KULLColors.Warning)Operation cancelled.$($KULLColors.Reset)"; Start-Sleep 1
                        continue
                    }
                    foreach ($m in $installedModules) {
                        Manage-Module -ModuleName $m.Name -Action Uninstall
                    }
                    Write-Centered "$($KULLColors.Success)All installed modules uninstalled. Press any key to continue...$($KULLColors.Reset)"
                    [void][System.Console]::ReadKey($true)
                    break
                } elseif ($sel -match '^\d+$' -and ($sel -as [int]) -ge 1 -and ($sel -as [int]) -le $installedModules.Count) {
                    $moduleToUninstall = $installedModules[($sel -as [int]) - 1]
                    Manage-Module -ModuleName $moduleToUninstall.Name -Action Uninstall
                    Write-Centered "$($KULLColors.Muted)Press any key to continue...$($KULLColors.Reset)"
                    [void][System.Console]::ReadKey($true)
                    # Refresh the list of installed modules
                    $installedModules = $modules | Where-Object { Get-Module -ListAvailable -Name $_.Name -ErrorAction SilentlyContinue }
                    if ($installedModules.Count -eq 0) {
                        Write-Centered "$($KULLColors.Success)All modules have been uninstalled.$($KULLColors.Reset)"; Start-Sleep 2
                        break
                    }
                } else {
                    Write-Centered "$($KULLColors.Failure)Invalid selection$($KULLColors.Reset)"; Start-Sleep 1
                }
            } while ($true)
        }
        'C' {
            [array]$list = $modules | Where-Object ConnectCmdletString
            if ($list.Count -eq 0) {
                Clear-Host
                Show-Header
                Write-Centered "$($KULLColors.Warning)No connectable modules configured.$($KULLColors.Reset)"; Start-Sleep 2
                continue
            }

            do {
                Clear-Host
                Show-Header
                Write-Host
                Write-Centered "$($KULLColors.Header)--- Connect to Services ---$($KULLColors.Reset)"
                Write-Host
                $keyColor = $KULLColors.MenuKey
                $reset = $KULLColors.Reset

                for ($i = 0; $i -lt $list.Count; $i++) {
                    $module = $list[$i]
                    $statusColor = $KULLColors.MenuText
                    $statusText = ""

                    if ($Global:AttemptedConnections.ContainsKey($module.Name)) {
                        if ($Global:AttemptedConnections[$module.Name].Success) {
                            $statusText = "$($KULLColors.Connected) (Connected)$reset"
                        } else {
                            $statusText = "$($KULLColors.FailedConnect) (Connection Failed)$reset"
                        }
                    }
                    $menuText = "$($KULLColors.MenuText)Connect to $($module.Description)$statusText"
                    Write-Centered "${keyColor}$($i + 1))$reset $menuText"
                }
                Write-Host

                $sel = Read-Host "$($KULLColors.Muted)Select a service to connect to, or press [Enter] to go back$($KULLColors.Reset)"

                if ($sel -eq '') { break }

                if ($sel -match '^\d+$' -and ($sel -as [int]) -ge 1 -and ($sel -as [int]) -le $list.Count) {
                    Invoke-ModuleConnect -ModuleToConnect $list[($sel - 1)]
                    Write-Centered "$($KULLColors.Muted)Press any key to continue...$($KULLColors.Reset)"
                    [void][System.Console]::ReadKey($true)
                } else {
                    Write-Centered "$($KULLColors.Failure)Invalid selection$($KULLColors.Reset)"; Start-Sleep 1
                }
            } while ($true)
        }
        'Q' {
            Clear-Host
            Write-Centered "$($KULLColors.Success)Goodbye!$($KULLColors.Reset)"
            Start-Sleep 1
            $exitLoop = $true
        }
        default {
        }
    }
} while (-not $exitLoop)