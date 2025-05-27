<#
.SYNOPSIS
Sets the PreferredDataLocation for a user in Microsoft 365

.DESCRIPTION
This PowerShell GUI tool uses Microsoft Graph to update a user’s PreferredDataLocation.
Supports UAC elevation, logging, and GUI-based input.

.AUTHOR
H4L-MK4

.VERSION
1.1

#>

# --- CONFIGURATION & LOG SETUP ---
$ErrorActionPreference = 'Stop'
$logDir       = 'C:\LOGS'
$timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'

# Ensure log directory exists
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    Write-Host "Created log directory: $logDir"
} else {
    Write-Host "Log directory already exists: $logDir"
}

# Start transcript to temp log in final directory
$tempLog = Join-Path $logDir "temp_$timestamp.log"
Start-Transcript -Path $tempLog -Force
Write-Host "Transcript started at $tempLog"

# --- ELEVATION CHECK & RESTART ---
Write-Host "Checking for elevation..."
$win = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $win.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Not elevated—relaunching as administrator..."
    $script = $MyInvocation.MyCommand.Definition
    Start-Process -FilePath powershell.exe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script`"" `
        -Verb RunAs
    Stop-Transcript
    exit
}
Write-Host "Running elevated as Administrator."

# --- LOAD GUI ASSEMBLIES ---
Write-Host "Loading GUI assemblies..."
try {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic
    Write-Host "GUI assemblies loaded."
}
catch {
    Write-Host "ERROR loading GUI assemblies: $($_.Exception.Message)"
    Stop-Transcript; exit 1
}

# --- ENSURE GRAPH MODULES ---
function Ensure-GraphModule { param($name)
    if (-not (Get-Module -ListAvailable -Name $name)) {
        Write-Host "Installing $name..."
        Install-Module $name -Scope AllUsers -Force -AllowClobber
    }
    Import-Module $name -ErrorAction Stop
    Write-Host "Module loaded: $name"
}
Write-Host "Ensuring Microsoft.Graph modules..."
Ensure-GraphModule -name 'Microsoft.Graph.Authentication'
Ensure-GraphModule -name 'Microsoft.Graph.Users'

# --- PROMPT FOR USER EMAIL ---
Write-Host "Prompting for user e-mail..."
$userEmail = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the e-mail address of the user:",
    "User E-mail",""
)
if ([string]::IsNullOrWhiteSpace($userEmail)) {
    Write-Host "No e-mail entered—exiting."
    [System.Windows.Forms.MessageBox]::Show("No e-mail entered. Exiting.","Cancelled",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
    Stop-Transcript; exit
}
Write-Host "User e-mail: $userEmail"

# --- PROMPT FOR PREFERRED DATA LOCATION ---
$locations  = @('CHE','IND','USA')
$promptText = @"
Enter the PreferredDataLocation code (CHE, IND, USA):
  CHE – Switzerland
  IND – India
  USA – United States
"@.Trim()
Write-Host "Prompting for PreferredDataLocation..."
$pdl = [Microsoft.VisualBasic.Interaction]::InputBox($promptText, 'PreferredDataLocation', 'CHE').Trim().ToUpper()
if (-not $locations.Contains($pdl)) {
    Write-Host "Invalid PDL entry: $pdl—exiting."
    [System.Windows.Forms.MessageBox]::Show("Invalid entry '$pdl'. Exiting.","Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    Stop-Transcript; exit 1
}
Write-Host "PreferredDataLocation selected: $pdl"

# --- CONNECT & UPDATE VIA GRAPH ---
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes 'User.ReadWrite.All','Group.ReadWrite.All'
Write-Host "Connected to Graph."
Try {
    Write-Host "Updating user $userEmail with PDL $pdl..."
    Update-MgUser -UserId $userEmail -PreferredDataLocation $pdl -ErrorAction Stop
    Write-Host "Update successful."
    [System.Windows.Forms.MessageBox]::Show(
        "PDL set to '$pdl' for '$userEmail'.","Success",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
}
Catch {
    Write-Host "Update failed: $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to update: $($_.Exception.Message)","Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    Stop-Transcript; exit 1
}

# --- CLEANUP & FINALIZE LOG ---
Stop-Transcript
Write-Host "Transcript stopped."

# Rename temp log to include safe e-mail
$safeEmail = $userEmail -replace '[^a-zA-Z0-9]', '_'
$finalLog = Join-Path $logDir "${safeEmail}_${timestamp}.log"
Rename-Item -Path $tempLog -NewName "${safeEmail}_${timestamp}.log" -Force
Write-Host "Log file renamed to: $finalLog"

# Open folder and select
Start-Process explorer.exe -ArgumentList "/select,`"$finalLog`""
exit
