# --- Make sure Powershell is running in 64 bit mode ----
Param([switch]$Is64Bit = $false)

Function Restart-As64BitProcess { 
    If ([System.Environment]::Is64BitProcess) { Write-Host "Already running as 64-bit process."; return } 
    $Invocation = $($MyInvocation.PSCommandPath)
    if ($Invocation -eq $null) { Write-Warning "Could not determine script path for relaunch."; return }
    Write-Warning "Attempting to restart script as 64-bit process..."
    $sysNativePath = $psHome.ToLower().Replace("syswow64", "sysnative")
    $ArgumentList = "-ExecutionPolicy Bypass -File `"$Invocation`" -Is64Bit" # Corrected -ex to -ExecutionPolicy and quoted $Invocation
    Start-Process "$sysNativePath\powershell.exe" -ArgumentList $ArgumentList -WindowStyle Hidden -Wait 
    exit # Exit the 32-bit instance after launching the 64-bit one
}

Restart-As64BitProcess


# --- Configuration ---
$oldOrgIdToOffboard = "A00000-0000-0000-0000-000000000" # Org ID of the tenant to offboard from find it here: 
$logDirectory = "C:\ProgramData\COMPANYNAME\DefenderOffOnboard"
$logFileName = "OffOnboardDefender_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$logFilePath = Join-Path -Path $logDirectory -ChildPath $logFileName
$flagFileName = "InstallOffOnBoardV1.flag" # User updated this
$flagFilePath = Join-Path -Path $logDirectory -ChildPath $flagFileName

# Name of your CMD script (ensure this is in the same folder as this PS1 script)
$offboardCmdScriptName = "NAMEOFCMD.cmd" 
$offboardCmdPath = Join-Path -Path $PSScriptRoot -ChildPath $offboardCmdScriptName

$onboardCmdScriptName = "NAMEOFONBOARDCMD.cmd" # Onboarding script
$onboardCmdPath = Join-Path -Path $PSScriptRoot -ChildPath $onboardCmdScriptName

# Name of your SharePoint Logger script (ensure this is in the same folder as this PS1 script)
# Removed this part since it requires complicated app registration.
#$sharePointLoggerScriptName = "SharePointLogger.ps1"
#$sharePointLoggerPath = Join-Path -Path $PSScriptRoot -ChildPath $sharePointLoggerScriptName

# --- Script Execution ---
try {
    # 1. Create Log Directory if it doesn't exist
    if (-not (Test-Path -Path $logDirectory -PathType Container)) {
        try {
            New-Item -Path $logDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Host "Log directory created: $logDirectory"
        }
        catch {
            Write-Error "Failed to create log directory: $logDirectory. Error: $($_.Exception.Message)"
        }
    }

    # 2. Start Transcript Logging
    try {
        Start-Transcript -Path $logFilePath -Force -ErrorAction Stop
        Write-Host "Transcript logging started to: $logFilePath"
    }
    catch {
        Write-Error "CRITICAL: Failed to start transcript logging to $logFilePath. Error: $($_.Exception.Message)"
    }

    Write-Host "Script execution started at $(Get-Date)"
    Write-Host "Running as user: $(whoami)"
    Write-Host "Is 64-bit PowerShell Process: $([System.Environment]::Is64BitProcess)"
    Write-Host "PSScriptRoot resolved to: $PSScriptRoot"
    Write-Host "Offboarding CMD script path: $offboardCmdPath"
    Write-Host "Onboarding CMD script path: $onboardCmdPath"
    #Write-Host "SharePoint Logger script path: $sharePointLoggerPath"

    # Removed this part since this is not needed for the on-off board
    ## 3. Dot-source SharePoint Logger (Your existing code for this) 
    #if (Test-Path -Path $sharePointLoggerPath) {
    #    try {
    #        . $sharePointLoggerPath 
    #        Write-Host "Successfully dot-sourced SharePointLogger.ps1"
    #        # Example: Log-ToSharePoint -Message "Defender offboarding/onboarding script started on $(hostname)." -Status "Information" 
    #    }
    #    catch { Write-Warning "Failed to dot-source SharePointLogger.ps1. Error: $($_.Exception.Message)" }
    #}
    #else { Write-Warning "SharePointLogger.ps1 not found. SharePoint logging skipped." }

    # 4. Defender Org ID Detection (Your existing code for this)
    $currentOrgId = $null
    $offboardNeeded = $false
    $oldErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue' 
    try {
        $currentOrgId = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status" -Name "OrgId" -ErrorAction Stop
    }
    catch { Write-Warning "Could not retrieve current Defender OrgId. Path or value might not exist." }
    $ErrorActionPreference = $oldErrorActionPreference 

    if ($currentOrgId) {
        Write-Host "Current Defender OrgId found: $currentOrgId"
        if ($currentOrgId -eq $oldOrgIdToOffboard) {
            Write-Host "Device is onboarded to the old Org ID ($oldOrgIdToOffboard). Offboarding required."
            $offboardNeeded = $true
        } else {
            Write-Host "Device is NOT onboarded to the target old Org ID. Current Org ID: $currentOrgId."
        }
    } else {
        Write-Host "No current Defender OrgId could be determined."
    }

    # --- ADDED ---
    $offboardingAttemptedAndSucceeded = $false 
    # --- END ADDED ---

    # 5. Execute Offboarding CMD if needed
    if ($offboardNeeded) {
        Write-Host "Attempting to run Defender Offboarding Script: $offboardCmdPath"
        if (Test-Path -Path $offboardCmdPath) {
            try {
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$offboardCmdPath`"" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                if ($process.ExitCode -eq 0) {
                    Write-Host "Offboarding CMD script executed successfully. Exit Code: $($process.ExitCode)"
                    $offboardingAttemptedAndSucceeded = $true # <-- SET FLAG HERE
                    # Example: Log-ToSharePoint -Message "Offboarding from $oldOrgIdToOffboard SUCCEEDED." -Status "Success"
                } else {
                    Write-Warning "Offboarding CMD script completed with non-zero Exit Code: $($process.ExitCode)."
                    # Example: Log-ToSharePoint -Message "Offboarding from $oldOrgIdToOffboard FAILED. Exit Code: $($process.ExitCode)." -Status "Error"
                }
            }
            catch {
                Write-Error "PowerShell error executing offboarding CMD '$offboardCmdPath'. Error: $($_.Exception.Message)"
                # Example: Log-ToSharePoint -Message "PowerShell error executing offboarding CMD: $($_.Exception.Message)" -Status "Error"
            }
        }
        else {
            Write-Error "Offboarding CMD script not found at: $offboardCmdPath."
            # Example: Log-ToSharePoint -Message "Offboarding CMD script NOT FOUND at $offboardCmdPath." -Status "Error"
        }
    }
    else {
        Write-Host "Offboarding was not deemed necessary for Org ID $oldOrgIdToOffboard."
    }

    # 6. Execute Onboarding CMD if Offboarding was successful
    if ($offboardingAttemptedAndSucceeded) { # Only true if $offboardNeeded was true AND $process.ExitCode was 0
        Write-Host "Offboarding successful. Attempting to run Defender Onboarding Script: $onboardCmdPath"
        # Example: Log-ToSharePoint -Message "Attempting onboarding to new tenant." -Status "Information"
        if (Test-Path -Path $onboardCmdPath) {
            try {
                # It's good practice to wait a few seconds after offboarding before starting onboarding
                Write-Host "Pausing for 15 seconds before starting onboarding..."
                Start-Sleep -Seconds 15

                $onboardProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$onboardCmdPath`"" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                if ($onboardProcess.ExitCode -eq 0) {
                    Write-Host "Onboarding CMD script executed successfully. Exit Code: $($onboardProcess.ExitCode)"
                    # Example: Log-ToSharePoint -Message "Onboarding CMD script SUCCEEDED." -Status "Success"
                } else {
                    Write-Warning "Onboarding CMD script completed with non-zero Exit Code: $($onboardProcess.ExitCode)."
                    # Example: Log-ToSharePoint -Message "Onboarding CMD script FAILED. Exit Code: $($onboardProcess.ExitCode)." -Status "Error"
                }
            }
            catch {
                Write-Error "PowerShell error executing onboarding CMD '$onboardCmdPath'. Error: $($_.Exception.Message)"
                # Example: Log-ToSharePoint -Message "PowerShell error executing onboarding CMD: $($_.Exception.Message)" -Status "Error"
            }
        }
        else {
            Write-Error "Onboarding CMD script not found at: $onboardCmdPath."
            # Example: Log-ToSharePoint -Message "Onboarding CMD script NOT FOUND at $onboardCmdPath." -Status "Error"
        }
    } elseif ($offboardNeeded -and (-not $offboardingAttemptedAndSucceeded)) {
        Write-Warning "Onboarding will not be attempted because the preceding offboarding operation did not succeed."
        # Example: Log-ToSharePoint -Message "Onboarding skipped because offboarding from old tenant failed or was not executed." -Status "Warning"
    }


    # 7. Create Flag File (Your existing code, just renumbered section)
    try {
        Write-Host "Creating flag file: $flagFilePath"
        if (-not (Test-Path -Path (Split-Path $flagFilePath) -PathType Container)) {
            New-Item -Path (Split-Path $flagFilePath) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }
        New-Item -Path $flagFilePath -ItemType File -Force -ErrorAction Stop | Out-Null
        Set-Content -Path $flagFilePath -Value "Defender offboarding/onboarding script logic completed processing at $(Get-Date). Offboard needed: $offboardNeeded. Offboarding Succeeded: $offboardingAttemptedAndSucceeded." -ErrorAction Stop
        Write-Host "Flag file created successfully: $flagFilePath"
    }
    catch { Write-Error "Failed to create flag file '$flagFilePath'. Error: $($_.Exception.Message)" }

    Write-Host "Script execution finished at $(Get-Date)."

}
catch {
    $errorMessage = "An unexpected critical error occurred: $($_.Exception.Message)"
    Write-Error $errorMessage
    if ($_.ScriptStackTrace) { Write-Error "Script stack trace: $($_.ScriptStackTrace)"}
}
finally {
    try {
        Stop-Transcript -ErrorAction SilentlyContinue # SilentlyContinue in case it was already stopped or failed to start
        Write-Host "Transcript logging stopped." # This might not appear in transcript if it's already stopped.
    }
    catch {
        # This catch is for Stop-Transcript itself failing.
        Write-Warning "Error stopping transcript: $($_.Exception.Message)"
    }
    #Upload-SPLogFile -FolderPath "Defender Offboard\$($env:COMPUTERNAME)" -LocalFilePath "$logFilePath"
}
# End of script