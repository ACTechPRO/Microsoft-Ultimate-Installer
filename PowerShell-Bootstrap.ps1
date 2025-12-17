<#
.SYNOPSIS
    PowerShell 7 Bootstrap - Automatic Detection and Installation

.DESCRIPTION
    Ensures PowerShell 7+ is present. Installs if missing (winget or direct download),
    then relaunches the main installer with PowerShell 7 for proper UTF-8 support.

.NOTES
    Version: 1.0.1
    Compatibility: PowerShell 5.1+
#>

param(
    [switch]$Force,
    [switch]$SkipBootstrap
)

#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info'
    )
    $colors = @{ Info='Cyan'; Success='Green'; Warning='Yellow'; Error='Red' }
    Write-Host "[$Level] $Message" -ForegroundColor $colors[$Level]
}

function Test-PowerShell7Available {
    try {
        $pwsh = Get-Command -Name pwsh.exe -ErrorAction SilentlyContinue
        if ($pwsh) {
            $major = & $pwsh.Source -NoProfile -Command '$PSVersionTable.PSVersion.Major'
            return ([int]$major -ge 7)
        }
    } catch { return $false }
    return $false
}

function Install-PowerShell7 {
    Write-Status "PowerShell 7 not found. Installing now..." -Level 'Warning'

    $wingetAvailable = $null -ne (Get-Command -Name winget.exe -ErrorAction SilentlyContinue)
    if ($wingetAvailable) {
        Write-Status "Installing PowerShell 7 via winget..." -Level 'Info'
        try {
            & winget.exe install --id Microsoft.PowerShell --source winget --accept-source-agreements --accept-package-agreements --scope machine 2>&1 | Out-Null
            Write-Status "PowerShell 7 installed via winget." -Level 'Success'
            return $true
        } catch {
            Write-Status "Winget install failed, trying direct download..." -Level 'Warning'
        }
    }

    Write-Status "Installing PowerShell 7 via direct download..." -Level 'Info'
    try {
        $tempDir = Join-Path $env:TEMP 'PowerShell7_Install'
        if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force }
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        $installerUrl = 'https://github.com/PowerShell/PowerShell/releases/download/v7.4.1/PowerShell-7.4.1-win-x64.msi'
        $installerPath = Join-Path $tempDir 'PowerShell-7.4.1-win-x64.msi'
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        (New-Object System.Net.WebClient).DownloadFile($installerUrl, $installerPath)

        $msiArgs = @('/i', $installerPath, '/quiet', '/norestart', 'ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1', 'ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1')
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $msiArgs -Wait -PassThru
        if ($proc.ExitCode -eq 0) {
            Write-Status 'PowerShell 7 installed successfully.' -Level 'Success'
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            return $true
        }
        Write-Status "Installation failed with exit code $($proc.ExitCode)." -Level 'Error'
        return $false
    } catch {
        Write-Status "Direct download installation failed: $_" -Level 'Error'
        return $false
    }
}

function Start-MainInstaller {
    $scriptPath = Join-Path $PSScriptRoot 'Microsoft 365 Ultimate Installer.ps1'
    if (-not (Test-Path $scriptPath)) {
        Write-Status "Main installer script not found at: $scriptPath" -Level 'Error'
        exit 1
    }
    Write-Status 'Launching Microsoft 365 Ultimate Installer...' -Level 'Info'
    try {
        $procArgs = @('-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-File',$scriptPath,'-IsHidden')
        if ($Force) { $procArgs += '-Force' }
        & pwsh.exe @procArgs
        exit
    } catch {
        Write-Status "Failed to launch installer: $_" -Level 'Error'
        exit 1
    }
}

Write-Host ''
Write-Host '=== Microsoft 365 Ultimate Installer - Bootstrap ===' -ForegroundColor Cyan
Write-Host '      PowerShell 7 Dependency Check' -ForegroundColor Cyan
Write-Host ''

if ($SkipBootstrap) {
    Write-Status 'Bootstrap skipped (already running in PowerShell 7).' -Level 'Info'
    exit 0
}

if (Test-PowerShell7Available) {
    Write-Status 'PowerShell 7+ detected.' -Level 'Success'
    Start-MainInstaller
}
else {
    Write-Status 'PowerShell 7 not found.' -Level 'Warning'
    if (Install-PowerShell7) {
        Write-Status 'Relaunching with PowerShell 7...' -Level 'Info'
        Start-Sleep -Seconds 2
        $selfPath = Join-Path $PSScriptRoot 'PowerShell-Bootstrap.ps1'
        $relaunchArgs = @('-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-File',$selfPath)
        if ($Force) { $relaunchArgs += '-Force' }
        $relaunchArgs += '-SkipBootstrap'
        & pwsh.exe @relaunchArgs
        exit
    }
    else {
        Write-Status 'Failed to install PowerShell 7.' -Level 'Error'
        Write-Status 'Please install PowerShell 7 manually: https://github.com/PowerShell/PowerShell/releases' -Level 'Error'
        exit 1
    }
}
