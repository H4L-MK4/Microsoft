# Microsoft

Some Microsoft things.

## Installing ExchangeOnlineManagement for All Users

The script `Install-Office365Modules.ps1` installs modules for all users by
default. To quickly ensure the ExchangeOnlineManagement module is available for
every PowerShell version, you can run the following snippet directly:

```powershell
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Scope AllUsers -Force -AllowClobber
} else {
    Write-Host "ExchangeOnlineManagement is already installed."
}
```

This installs the module to `C:\Program Files\WindowsPowerShell\Modules`, which
is shared across all PowerShell versions on the system.
