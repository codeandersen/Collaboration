# Requires: Run PowerShell as Administrator
# Purpose : Make SCHANNEL and .NET TLS settings explicit 
# Notes   : Disables TLS 1.0/1.1, Enables TLS 1.2, sets .NET strong crypto defaults
# Reboot  : Required after applying

$ErrorActionPreference = "Stop"

$script:LogFile = Join-Path -Path $PSScriptRoot -ChildPath ("Set-ExchangeTls-HealthCheckerFix_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    try {
        Add-Content -Path $script:LogFile -Value $line -ErrorAction Stop
    } catch {
        Write-Host "Failed to write to log file '$script:LogFile': $($_.Exception.Message)" -ForegroundColor Red
    }

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Gray }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }
}

function Assert-Administrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log -Level 'ERROR' -Message "This script must be run as Administrator. Exiting."
        throw "Not running elevated"
    }
}

function Set-RegDword {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][int]$Value
    )

    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
        }

        New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force -ErrorAction Stop | Out-Null
        Write-Log -Message "Set DWORD: $Path\\$Name = $Value"
    } catch {
        Write-Log -Level 'ERROR' -Message "Failed to set DWORD: $Path\\$Name = $Value. $($_.Exception.Message)"
        throw
    }
}

function Set-TlsProtocol {
    param(
        [Parameter(Mandatory=$true)][ValidateSet("TLS 1.0","TLS 1.1","TLS 1.2")][string]$Protocol,
        [Parameter(Mandatory=$true)][ValidateSet("Client","Server")][string]$Role,
        [Parameter(Mandatory=$true)][bool]$Enable
    )

    $base = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$Protocol\$Role"

    if ($Enable) {
        # Enabled=1, DisabledByDefault=0
        Set-RegDword -Path $base -Name "Enabled" -Value 1
        Set-RegDword -Path $base -Name "DisabledByDefault" -Value 0
    } else {
        # Enabled=0, DisabledByDefault=1
        Set-RegDword -Path $base -Name "Enabled" -Value 0
        Set-RegDword -Path $base -Name "DisabledByDefault" -Value 1
    }
}

try {
    Assert-Administrator
    Write-Log -Message "Starting TLS/.NET registry configuration"
    Write-Log -Message "Log file: $script:LogFile"

    Write-Log -Message "Setting SCHANNEL protocol registry values..."

    # Disable TLS 1.0 and TLS 1.1 (Client + Server)
    Set-TlsProtocol -Protocol "TLS 1.0" -Role "Client" -Enable:$false
    Set-TlsProtocol -Protocol "TLS 1.0" -Role "Server" -Enable:$false
    Set-TlsProtocol -Protocol "TLS 1.1" -Role "Client" -Enable:$false
    Set-TlsProtocol -Protocol "TLS 1.1" -Role "Server" -Enable:$false

    # Enable TLS 1.2 (Client + Server)
    Set-TlsProtocol -Protocol "TLS 1.2" -Role "Client" -Enable:$true
    Set-TlsProtocol -Protocol "TLS 1.2" -Role "Server" -Enable:$true

    Write-Log -Message "Setting .NET Framework TLS defaults (SystemDefaultTlsVersions + SchUseStrongCrypto)..."

    $netPaths = @(
        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319",
        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727"
    )

    foreach ($p in $netPaths) {
        Set-RegDword -Path $p -Name "SystemDefaultTlsVersions" -Value 1
        Set-RegDword -Path $p -Name "SchUseStrongCrypto"       -Value 1
    }

    Write-Log -Message "Done. A REBOOT is required for SCHANNEL changes to fully apply." -Level 'WARN'
} catch {
    Write-Log -Level 'ERROR' -Message "Script failed: $($_.Exception.Message)"
    throw
}
