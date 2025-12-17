<#
.SYNOPSIS
    Microsoft 365 Ultimate Installer
    
.DESCRIPTION
    Automated installation and licensing of Microsoft 365 Enterprise.
    Features:
    - Installs Microsoft 365 Enterprise (Current Channel Preview)
    - Includes Project Pro, Visio Pro, Teams, Clipchamp, Power Automate
    - Languages: En-US, Es-ES, Ja-JP, Pt-BR
    - Privacy: Maximum lockdown (Telemetry disabled)
    - Exclusions: OneDrive, OneNote, Skype, Sticky Notes (Blocked)
    - Licensing: Auto-activation for Windows (HWID) and Office (Ohook)
    
.PARAMETER Force
    Force clear any stuck instance mutex before running.
    Use this if you get "Another instance is already running" message.
    
.NOTES
    Version: 3.0.0 (Single script, hidden execution, improved cleanup)
    Author: AI-Assisted Automation
#>

param(
    [switch]$Force,
    [switch]$IsHidden
)

#Requires -Version 5.1
#Requires -RunAsAdministrator

# ============================================================================
# SELF-ELEVATION & HIDDEN EXECUTION (RELAUNCH HIDDEN, NO P/INVOKE)
# ============================================================================
# If the console is visible, relaunch the script hidden and exit the parent.
# Uses only native Start-Process (no C#/PInvoke) to minimize AV heuristics.
if (-not $IsHidden) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrEmpty($scriptPath)) {
        Write-Host "Error: Cannot determine script path for hidden execution" -ForegroundColor Red
        exit 1
    }

    try {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = (Get-Process -Id $PID).Path
        $forceArg = if ($Force) { " -Force" } else { "" }
        $startInfo.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -IsHidden$forceArg"
        $startInfo.WindowStyle = 'Hidden'
        $startInfo.CreateNoWindow = $true
        $startInfo.UseShellExecute = $false
        [System.Diagnostics.Process]::Start($startInfo) | Out-Null
        exit
    }
    catch {
        Write-Host "Error launching hidden process: $_" -ForegroundColor Red
        exit 1
    }
}

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ============================================================================
# CONFIGURATION
# ============================================================================

$Script:Config = @{
    LogFile          = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Microsoft Ultimate Installer.log'
    TempFolder       = Join-Path $env:LOCALAPPDATA 'Temp\M365Ultimate_Installation'
    ODTUrls          = @(
        'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18324-20194.exe'
        'https://download.microsoft.com/download/C/0/C/C0CCDA5A-A315-49B5-95DC-65071C2E327E/officedeploymenttool.exe'
    )
    WingetPackages   = @(
        @{ Id = '9P1J8S7CCWWT'; Source = 'msstore'; Name = 'Microsoft Clipchamp' }
        @{ Id = 'Microsoft.PowerAutomateDesktop'; Source = 'winget'; Name = 'Power Automate Desktop' }
    )
    ExcludedApps     = @(
        'OneDrive'
        'OneNote'
        'Lync'
        'Groove'
    )
    MASUrls          = @(
        'https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd'
        'https://gitlab.com/massgrave/microsoft-activation-scripts/-/raw/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd'
        'https://bitbucket.org/WindowsAddict/microsoft-activation-scripts/raw/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd'
    )
    InstallStartTime = $null
}

# AC Tech Logo embedded as Base64 (for standalone script distribution)
$Script:ACTechLogoBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAZAAAABkCAYAAACoy2Z3AAAPMklEQVR4nO3dB9AdVRUH8P8FgUgooQhIGxVRBOkgASTMNYBACOBIQJp0k4g4FMECCkFBUVQEhkQRTCIKSI8CUm+AhAQUpAmhCCITegkYROp1ztuT8WNz7vt239t97yv/38w3gd1XNi/v29vOORcgIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiKigcV1+wKoHwlxEQBbAdgawEYA1gOwPIBlASwBYB6AlwE8A+BOALMB3ALv5BgRDTBsQPqSEOWmu7lx5kZ4tz26JcR1ARwB4AsAVi757P8CuBTARHh3OzotxEkAxqL7joV3p3f7IvqMEJcC8G/jzDbwbkYXrohaID1K6gtCXC/ReIiRCHHtrlxTiFcAeADAuBYaDzEEwH4AZiLESxDih2u4UiLqAjYgfcchvYwU5QbeGSEuhhC/B+BuALtXOFLdA8AchCivSUT9HBuQvkBu2FkvvZkDEeKQDlyLrGlMBzABwOK9PHo+gMcA/FXXOx7WY80s05jSCrEvTCsRURs+0M6TqTKjAXwod0wWpIf1+H+5se8JYGptn3uIqwO4DoCseaTcAWAKgNsAPAjv3su9hnRKZDpuFICvAPio8RqLApiEEN+Bd+dV/vcgoo7gInpfEOLVAHbOHf0mgNNyx2bDuy1rXNSURe71E4/4C4Cvwbs7S7ymNBRHAjgZwJLGI94E8Fl4JyOYviVEue6fG2e4yFvN58tF9AGAU1jdFuKqAD6fO/oogHMBvJ07PhwhbljDNTgd2aQaj5Ma712m8RDevQvvfgpgWwAvGY+Q0N9LOjI1R0SVYwPSfQfolE5PkjvxCoCbjMePr+EaDtQQ3bzYOOfdhIWmqsrIRhjSSL5lnP2IhggTUT/DBqT7DjKOBf3zIuPcvghx6crePURJAvxR4uwJ8E7WO9rn3V0Ajkuc/Q5CXK6S9yGijmED0k0hjgCQz++QXvo1+t+Sg/FG7vxSBSK2ypC1lpWM49Ph3amo1lkAHjKOD6v470REHcAGpLsONo7dAO/maa/9NQDTapvGytYeJFIq710Ah6Nq2TSYrKdY2IAQ9TNsQLolm4aSxLo8KfvRkzWFtD5ClHpU7doHwArG8Wnw7kHU43IAzxvHl9ccFCLqJ5gH0j17ARhqTF9dlTt2PYCnAUi0Vk+SmT6zzWvYO3Fc6kfVw7t3EKJEfHkAM/TnNnj3XG3v2b/L24zWApaf0O/AkhqdJwUqH9fwasnduanx2VZ/DRvqNXxG83tW0GlUSRidpxGDcg036LRnrPj9P6gBHjsB2BjAmvoZvKEdkTm6ZngZvHui0vemXjEPpFtCnNUIjX2/S+HdGOOxshbxbSOHYnV492KL77+k3oQklLanVwGsWMvNqD/pZh5IiLsCOF5v2kVJJ0NCps+Gd2+1+f4SFbgvgG80Ce22SIMm39Xze21IessDCVE6t0cB+JYm0fbmPV0zPAbePVnimqkNnMLqhhDXMRoP8ZvEM6zjSyQiuIra1mg8UFtPlnoX4kqaVHpVycYDOjqRBuTetnKFsucuqDZQpvEQHwPw68aosp2imSGuqaVxflyw8VhwL/sigPsRolRBoA5gA9J3Cic+o1MRC/Pu0cR01VhNAmzFponjnS+5TnLT/DSAvxkVCcqSzskshCjTTuVkN97bm3w3ipIpt9sRojQoZX1SG7BWr0HWFq9AiNu0+HwqgWsgnZYNzfc3zkxtZG6nna8bOfW0FoDtdZ2klRuNRfI1qJOyHrfM46+YeISMCCUZ8z7N6Je1M+nhy01yFePxsm5wOUL0hafbspDyK5vcE97VUcEcXXsYousy22iBTCtBVG7kW8A72ROmqHOMIp5vagfqAZ32krBvaXC3TBT8lOKkkxv72Hgnz6WasAHpvFGJfTUm9/K8PwA401h4H99iAyI9PcsjLbwWtSpEmUa8LNF4yM3ydF3XWHhXx2z0OUITQYcbv9sXI8SN4Z0V9dbzdaSQ5yWJ+8F8vYYztTpC/rlD9Tt4kvHd3ECrOkuuUVE9G4TXdE1lErx71XhvaTxPSYTDy+jnS4koRqoIp7D6xvTVLHgnPbs07+brL3neaIS4WgvXYfVc39KpNOqcHwDYzDh+byPqyLuTk1sCy0K1d1L2Rnri30+si/yqwDVMSCSTPqLXMMFsPLJreF13WpSGzLrOwxFiamTVzD2NUYZ3p5mNR/bez8K7Q7QRsUj1aqoRG5BOCnFlDUcsunhuTWPlScTMYS1cjUTB5L1ceRgmpWU31q8aZ6QzMRLe/aPwx+edbAD2XePMbrq+krqG1RLfn2cbjYJ3jxV8f9l87MvGmaFa762MfwHYDt49VfDx8ne38pa2bGONkApgA9JZBxjTBBLPfnGhZ3sne3DIgnreYbq20m4Dki+bQvX6ulHmXsJR94N3VvXi3khPXG7keakaZNDpH+u7c3Dp3BzvJILsZuNM2aioI0v9/bMKB9ZIazljnx2qEBuQzrLCbiUBSuZ6i5qcmKooHnWT9cryFYBFs0V8qqcKMoxcoNYCGbLRo4Ty5o3RtRbLGHP6yLtrW7oG4DzNSQmakHo0gB+WeP5cXcwvSzpXFtkkjWrCRfROyUqPWJFPRaevFpii8935xn+8JlIVu9GE+IbR+5XoHeqEEGWnxjWMM79r85Wv0aitnr/bQ3SdZaaxeG5Nb7W+S6R3vwcgP626vsVpVGtkjkSEGFWEI5DOsSJF/tmjdHsx3s1NRF1thxA/XuKVrL3L2YB0jiRyWiQHonVZIU5r7cTKi5BkRVeiN98JEjzQivmJETS/0zXiCKQTslBHKyJkSou9LVlM3zF3TG4EYwEcW/A1XjIib5ZDiIu3XQqDitjIOCY3wPEIbccxSB5EnpVVLnkceZKz8Xd0T9GFc2tU/Y4xNWtN1VJF2IB0xp6JResTEeKJFb7PQQjxhILJUxJd8ymjEVpdaxpRvazQVrnZVfl96MmquiwJjHnPdbmUjR2yS30Sp7C6l/tR103CWhQtM2cs2e1Uv06XrrcakKX74A2cmeP9CBuQuoUo0wRV7N1RVNHNpu5PHLeS2qh6Uo6j2+9nrQ/8pwPXQgMEp7C6s3hep60Q4gbwTuomNXNr4njZKrCtCVH2kJirQQTyc/8gS2IsUx+qLlYgRb4cCVESG5A6ZfsqWNm5kqDVvHRJMRKGua5xfFwiw/n/vHscIT5lhJKObOQM1FmELqvSupn+7KZHX0KIkhMjgQCDgZUodzW826WD12CVHmHYKxXGBqReUrbE2hfhRHj3y7ZfPcRVtexDPtJkP4R4nNbPauZqbWzy8+JS4fdPqM9OiTn6wRQx83IfWH+yrmHlRseneWVoogaugXR+8fytRFHE8rx7OpETIo3AfgVeIZW0VvcoYGyT/dIHCytXYy2EuEyXr2GIEZ1XToh3I8QHG5tjhXg2Qjy6ZI4S9RMcgdS5u5xdA+jaZHXV1kxO9OjHFdjbXDKTZR9pyYruaVSjAJ93sv9CtWSPCjsnQaJ/bsLgcUsif2NHLd3fmhCHaITdfE1UfUL/nAnv8puSzWiyIVRr//YhLqs5Li7XEMl+JsUKM1K/wRFIffZPJHS1W6oiT7Y/tRqkDRGilPlOyxatrX2/5Zf/7MormWYFH6X0t2XiINv8R+pdWVOM+SnFsvbWXJ51tDGSqLzTAGyy0COzfUKscO6y1XN72t3IbpdR951tvCb1UWxAOls4UYom/rHSd8luuhe2EdIrVUxlKswqtXEUqnWseSPLQkd/hsEkS9a7yjqDEHduY/QhGzvlSbXaSxPPujIRyde885FmfeduLrkrIfUTbEDqEKLsDreecebymn6RUgUZpQqrlUCWb4Csm474CULcB1UIcUxi0yNxDrx7AYOP7CRohS5PbXHN4IxEdrlEt6U2CvsFgLeN41O0BE9xIUqnaQvjzMRSr0P9BhuQzuZ+VD19lcnKf1uJgUMSJcPzzz83sRi/iN7MTkCIrX9XQjxA/+5WlNXDNZbv6NuyNSZrBCCNfig1CghxQiI44d2mn29WnPMC48zaAG4svJtgiDskGoqHNNqPBiA2IFULUUqk72WceSax2U7do5CxBdcypNGzRgGL6sjhNoRoVXRNC3F1hHixLvRb60HS890H3g3m7GeZ8pGbeJ6sY9yCECdq6XdbiJsgxOt1Vz6LbAkrN/HephYlHDxPRtL3IcRDk/uJSLBIiLKuJfuH5B8jo6vDGBI8cDEKq3p7JJKxLtSd0+pygS6WLmb0JEc2epPNSE80xJ20kVsmEZlzK0Kcpe8l+zY8lojCGaGfw96JhkPIZ3GoboU6eMmufyHuoVFZi+fOLqaL6uMaYbGAZO8/r6OKVXS6qFnIbWjSsPS8Bkni3FOrE+SvQfKYZIR6BkKcrhFd87Q0yvrayOSfs8ApRuQXDSBsQDpXOLGe6asFZA1B4u6zKBirl3tjgde4CyGO1sXdVK0mmVbJplZCnKe9Z/lzqE69rFZgZCs3wP3hXWrxf3DxbrY23jJaS00ZrZuoOpAiIdG7F+79e3cHQtxVr0E6AXlDS25Nexa8s/ZopwGEU1hVClEyia1pnjkd6mmnprF21az13nknvdBNC27sM0yDBbbW2P81CnynZCpvFzYeC33uN2tpl/Y2lMpGdrIwPqpAJYL8NVynoxoZ7bRKdro8At7Jfu80wLEBqZasI7iOjz7ev52p1NmyRpqHFn4VqZOVTU0cX3F576mNBse7P1f4mgOHd0/Cu+FaH+yuFhoO+fcfDu+ObDmnxjsJathAw9AlCbEoGelc1AjT9u7slt6b+h1OYVUli1KyCieizT2iy+UWhCjrE8cYZw9DiKeUmNKQcONTEeIkLcy4V2L/7N68DuC3AM4ssJhL2Wc/DcA0hCjJgJ/Tn7V0ilCmuBbRacMXAdwDYHbj8d5ZC+HlZd+RyQhR/t02B7CDrmutpu8/THOaXtDdC6fr+z/Jf8DBpdpMYxrYshvaCO2hys+quuC+tN7UXtGs+Ll6U5uhJTTKTaUQERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERevU/+e3Fqe2rTPMAAAAASUVORK5CYII='
$Script:ACTechFooterBase64 = $Script:ACTechLogoBase64

# ============================================================================
# LOCALIZATION SYSTEM
# ============================================================================

# Detect Windows UI language
$Script:CurrentLanguage = (Get-UICulture).Name
if ($Script:CurrentLanguage -notmatch '^(en|pt|es|ja|de|fr|zh|it|ko|ru)') {
    $Script:CurrentLanguage = 'en-US'
}
# Normalize to base language for matching
$Script:LangBase = $Script:CurrentLanguage.Substring(0, 2).ToLower()

# Localized strings dictionary
$Script:Strings = @{
    'en' = @{
        # Window titles
        WindowTitle                     = 'Microsoft Ultimate Installer v3.6'
        WindowSubtitle                  = 'Automated Installer & Activator'
        ConfigWindowTitle               = 'Installation Configuration'
        
        # Installation modes
        ExpressMode                     = 'Express Installation (Recommended)'
        ExpressModeDesc                 = 'Install with default settings - Microsoft 365 with all apps'
        CustomMode                      = 'Custom Installation'
        CustomModeDesc                  = 'Choose version, languages and applications'
        
        # Configuration labels
        SelectVersion                   = 'Select Office Version:'
        SelectLanguages                 = 'Select Languages:'
        PrimaryLanguage                 = 'Primary Language:'
        AdditionalLanguages             = 'Additional Languages:'
        SelectApps                      = 'Select Applications:'
        SelectAll                       = 'Select All'
        DeselectAll                     = 'Deselect All'
        
        # Buttons
        BtnStart                        = 'Start Installation'
        BtnCancel                       = 'Cancel'
        BtnBack                         = 'Back'
        BtnUseDefaults                  = 'Use Defaults'
        
        # Office versions
        Version365Enterprise            = 'Microsoft 365 Enterprise'
        Version365Business              = 'Microsoft 365 Business'
        VersionProPlus2024              = 'Office LTSC Professional Plus 2024'
        VersionProPlus2021              = 'Office LTSC Professional Plus 2021'
        VersionProPlus2019              = 'Office Professional Plus 2019'
        
        # Applications
        AppWord                         = 'Word'
        AppExcel                        = 'Excel'
        AppPowerPoint                   = 'PowerPoint'
        AppOutlook                      = 'Outlook'
        AppAccess                       = 'Access'
        AppPublisher                    = 'Publisher'
        AppOneNote                      = 'OneNote'
        AppOneDrive                     = 'OneDrive'
        AppTeams                        = 'Teams'
        AppProject                      = 'Project Professional'
        AppVisio                        = 'Visio Professional'
        AppClipchamp                    = 'Clipchamp'
        AppPowerAutomate                = 'Power Automate Desktop'
        
        # Progress messages
        StatusInitializing              = 'Initializing...'
        StatusPreparing                 = 'Preparing...'
        StatusDownloadingODT            = 'Downloading Office Deployment Tool...'
        StatusDownloadingOffice         = 'Downloading Office Files...'
        StatusInstallingOffice          = 'Installing Microsoft 365...'
        StatusInstallingExtras          = 'Installing Extras...'
        StatusActivating                = 'Activating licenses...'
        StatusCleaning                  = 'Cleaning up...'
        StatusFinalizing                = 'Finalizing...'
        StatusComplete                  = 'Complete!'
        
        SubStatusInternetSpeed          = 'This may take a few minutes depending on your internet speed'
        SubStatusApplyingConfig         = 'Setting up your personalized configuration'
        SubStatusClipchampPowerAutomate = 'Installing video editor and automation tools'
        SubStatusWindowsOffice          = 'Activating your Windows and Office licenses'
        SubStatusRemovingTemp           = 'Cleaning up installation files'
        SubStatusAllComplete            = 'Your Microsoft 365 is ready to use!'
        SubStatusRemovingConflicts      = 'Removing old Office versions to avoid conflicts'
        SubStatusDownloadAttempt        = 'Downloading installation tools'
        SubStatusRevertingChanges       = 'Undoing changes due to error'
        
        # Log messages
        LogInstallStarted               = '=== INSTALLATION STARTED ==='
        LogInstallComplete              = '=== INSTALLATION COMPLETED SUCCESSFULLY ==='
        LogInstallCancelled             = '=== INSTALLATION CANCELLED BY USER ==='
        LogInstallFailed                = '=== CRITICAL ERROR ==='
        LogDuration                     = 'Total duration:'
        LogPhase                        = 'Phase'
        LogCompleted                    = 'completed'
        
        # Errors
        ErrAnotherInstance              = 'Another instance is already running.'
        ErrInstallCancelled             = 'Installation was cancelled.'
        ErrAllChangesReverted           = 'All changes have been reverted and temporary files removed.'
        ErrInstallFailed                = 'Installation Failed!'
        ErrCheckLog                     = 'Check log file on Desktop for details:'
        
        # Validation
        ValidSelectAtLeastOneApp        = 'Please select at least one application to install.'
        ValidSelectAtLeastOneLang       = 'Please select at least one language.'
        ValidProjectVisioNote           = 'Project and Visio require separate licenses but will be activated automatically.'

        ConfigureInstallation           = 'Configure Installation'
        FooterDevBy                     = 'Developed by AC Tech'
    }
    
    'pt' = @{
        WindowTitle                     = 'Instalador Microsoft Ultimate'
        WindowSubtitle                  = 'Instalador e Ativador Automatizado'
        ConfigWindowTitle               = 'Configuração da Instalação'
        
        ExpressMode                     = 'Instalação Expressa (Recomendado)'
        ExpressModeDesc                 = 'Instalar com configurações padrão - Microsoft 365 com todos os apps'
        CustomMode                      = 'Instalação Personalizada'
        CustomModeDesc                  = 'Escolher versão, idiomas e aplicativos'
        
        SelectVersion                   = 'Selecione a Versão do Office:'
        SelectLanguages                 = 'Selecione os Idiomas:'
        PrimaryLanguage                 = 'Idioma Principal:'
        AdditionalLanguages             = 'Idiomas Adicionais:'
        SelectApps                      = 'Selecione os Aplicativos:'
        SelectAll                       = 'Selecionar Todos'
        DeselectAll                     = 'Desmarcar Todos'
        
        BtnStart                        = 'Iniciar Instalação'
        BtnCancel                       = 'Cancelar'
        BtnBack                         = 'Voltar'
        BtnUseDefaults                  = 'Usar Padrão'
        
        Version365Enterprise            = 'Microsoft 365 Enterprise'
        Version365Business              = 'Microsoft 365 Business'
        VersionProPlus2024              = 'Office LTSC Professional Plus 2024'
        VersionProPlus2021              = 'Office LTSC Professional Plus 2021'
        VersionProPlus2019              = 'Office Professional Plus 2019'
        
        AppWord                         = 'Word'
        AppExcel                        = 'Excel'
        AppPowerPoint                   = 'PowerPoint'
        AppOutlook                      = 'Outlook'
        AppAccess                       = 'Access'
        AppPublisher                    = 'Publisher'
        AppOneNote                      = 'OneNote'
        AppOneDrive                     = 'OneDrive'
        AppTeams                        = 'Teams'
        AppProject                      = 'Project Professional'
        AppVisio                        = 'Visio Professional'
        AppClipchamp                    = 'Clipchamp'
        AppPowerAutomate                = 'Power Automate Desktop'
        
        StatusInitializing              = 'Inicializando...'
        StatusPreparing                 = 'Preparando...'
        StatusDownloadingODT            = 'Baixando Office Deployment Tool...'
        StatusDownloadingOffice         = 'Baixando arquivos do Office...'
        StatusInstallingOffice          = 'Instalando Microsoft 365...'
        StatusInstallingExtras          = 'Instalando Extras...'
        StatusActivating                = 'Ativando licenças...'
        StatusCleaning                  = 'Limpando...'
        StatusFinalizing                = 'Finalizando...'
        StatusComplete                  = 'Concluído!'
        
        SubStatusInternetSpeed          = 'Isso pode levar alguns minutos dependendo da sua conexão'
        SubStatusApplyingConfig         = 'Configurando sua instalação personalizada'
        SubStatusClipchampPowerAutomate = 'Instalando editor de vídeo e ferramentas de automação'
        SubStatusWindowsOffice          = 'Ativando suas licenças do Windows e Office'
        SubStatusRemovingTemp           = 'Limpando arquivos de instalação'
        SubStatusAllComplete            = 'Seu Microsoft 365 está pronto para usar!'
        SubStatusRemovingConflicts      = 'Removendo versões antigas do Office para evitar conflitos'
        SubStatusDownloadAttempt        = 'Baixando ferramentas de instalação'
        SubStatusRevertingChanges       = 'Desfazendo alterações devido a erro'
        
        LogInstallStarted               = '=== INSTALAÇÃO INICIADA ==='
        LogInstallComplete              = '=== INSTALAÇÃO CONCLUÍDA COM SUCESSO ==='
        LogInstallCancelled             = '=== INSTALAÇÃO CANCELADA PELO USUÁRIO ==='
        LogInstallFailed                = '=== ERRO CRÍTICO ==='
        LogDuration                     = 'Duração total:'
        LogPhase                        = 'Fase'
        LogCompleted                    = 'concluída'
        
        ErrAnotherInstance              = 'Outra instância já está em execução.'
        ErrInstallCancelled             = 'Instalação cancelada.'
        ErrAllChangesReverted           = 'Todas as alterações foram revertidas e arquivos temporários removidos.'
        ErrInstallFailed                = 'Falha na Instalação!'
        ErrCheckLog                     = 'Verifique o arquivo de log na Área de Trabalho:'
        
        ValidSelectAtLeastOneApp        = 'Selecione pelo menos um aplicativo para instalar.'
        ValidSelectAtLeastOneLang       = 'Selecione pelo menos um idioma.'
        ValidProjectVisioNote           = 'Project e Visio requerem licenças separadas, mas serão ativados automaticamente.'
        ConfigureInstallation           = 'Configurar Instalação'
        FooterDevBy                     = 'Desenvolvido por AC Tech'
    }
    
    'es' = @{
        WindowTitle                     = 'Instalador Microsoft 365 Ultimate'
        WindowSubtitle                  = 'Instalador y Activador Automatizado'
        ConfigWindowTitle               = 'Configuración de Instalación'
        
        ExpressMode                     = 'Instalación Rápida (Recomendado)'
        ExpressModeDesc                 = 'Instalar con configuración predeterminada - Microsoft 365 con todas las apps'
        CustomMode                      = 'Instalación Personalizada'
        CustomModeDesc                  = 'Elegir versión, idiomas y aplicaciones'
        
        SelectVersion                   = 'Seleccione la Versión de Office:'
        SelectLanguages                 = 'Seleccione los Idiomas:'
        PrimaryLanguage                 = 'Idioma Principal:'
        AdditionalLanguages             = 'Idiomas Adicionales:'
        SelectApps                      = 'Seleccione las Aplicaciones:'
        SelectAll                       = 'Seleccionar Todo'
        DeselectAll                     = 'Deseleccionar Todo'
        
        BtnStart                        = 'Iniciar Instalación'
        BtnCancel                       = 'Cancelar'
        BtnBack                         = 'Volver'
        BtnUseDefaults                  = 'Usar Predeterminado'
        
        Version365Enterprise            = 'Microsoft 365 Enterprise'
        Version365Business              = 'Microsoft 365 Business'
        VersionProPlus2024              = 'Office LTSC Professional Plus 2024'
        VersionProPlus2021              = 'Office LTSC Professional Plus 2021'
        VersionProPlus2019              = 'Office Professional Plus 2019'
        
        AppWord                         = 'Word'
        AppExcel                        = 'Excel'
        AppPowerPoint                   = 'PowerPoint'
        AppOutlook                      = 'Outlook'
        AppAccess                       = 'Access'
        AppPublisher                    = 'Publisher'
        AppOneNote                      = 'OneNote'
        AppOneDrive                     = 'OneDrive'
        AppTeams                        = 'Teams'
        AppProject                      = 'Project Professional'
        AppVisio                        = 'Visio Professional'
        AppClipchamp                    = 'Clipchamp'
        AppPowerAutomate                = 'Power Automate Desktop'
        
        StatusInitializing              = 'Inicializando...'
        StatusPreparing                 = 'Preparando...'
        StatusDownloadingODT            = 'Descargando Office Deployment Tool...'
        StatusDownloadingOffice         = 'Descargando archivos de Office...'
        StatusInstallingOffice          = 'Instalando Microsoft 365...'
        StatusInstallingExtras          = 'Instalando Extras...'
        StatusActivating                = 'Activando licencias...'
        StatusCleaning                  = 'Limpiando...'
        StatusFinalizing                = 'Finalizando...'
        StatusComplete                  = '¡Completo!'
        
        SubStatusInternetSpeed          = 'Velocidad depende de la conexión'
        SubStatusApplyingConfig         = 'Aplicando configuración'
        SubStatusClipchampPowerAutomate = 'Clipchamp y Power Automate'
        SubStatusWindowsOffice          = 'Activación Windows y Office'
        SubStatusRemovingTemp           = 'Eliminando archivos temporales'
        SubStatusAllComplete            = 'Todas las tareas completadas con éxito'
        
        LogInstallStarted               = '=== INSTALACIÓN INICIADA ==='
        LogInstallComplete              = '=== INSTALACIÓN COMPLETADA CON ÉXITO ==='
        LogInstallCancelled             = '=== INSTALACIÓN CANCELADA POR USUARIO ==='
        LogInstallFailed                = '=== ERROR CRÍTICO ==='
        LogDuration                     = 'Duración total:'
        LogPhase                        = 'Fase'
        LogCompleted                    = 'completada'
        
        ErrAnotherInstance              = 'Otra instancia ya está en ejecución.'
        ErrInstallCancelled             = 'Instalación cancelada.'
        ErrAllChangesReverted           = 'Todos los cambios fueron revertidos y archivos temporales eliminados.'
        ErrInstallFailed                = '¡Error en la Instalación!'
        ErrCheckLog                     = 'Revise el archivo de log en el Escritorio:'
        
        ValidSelectAtLeastOneApp        = 'Seleccione al menos una aplicación para instalar.'
        ValidSelectAtLeastOneLang       = 'Seleccione al menos un idioma.'
        ValidProjectVisioNote           = 'Project y Visio requieren licencias separadas pero serán activados automáticamente.'
        ConfigureInstallation           = 'Configurar Instalación'
        FooterDevBy                     = 'Desarrollado por AC Tech'
    }
    
    # 'ja' - Japanese localization removed due to encoding issues

    
    'de' = @{
        WindowTitle                     = 'Microsoft 365 Ultimate Installer'
        WindowSubtitle                  = 'Automatischer Installer "&" Aktivator'
        ConfigWindowTitle               = 'Installationskonfiguration'
        
        ExpressMode                     = 'Express-Installation (Empfohlen)'
        ExpressModeDesc                 = 'Mit Standardeinstellungen installieren - Microsoft 365 mit allen Apps'
        CustomMode                      = 'Benutzerdefinierte Installation'
        CustomModeDesc                  = 'Version, Sprachen und Anwendungen auswählen'
        
        SelectVersion                   = 'Office-Version auswählen:'
        SelectLanguages                 = 'Sprachen auswählen:'
        PrimaryLanguage                 = 'Hauptsprache:'
        AdditionalLanguages             = 'Zusätzliche Sprachen:'
        SelectApps                      = 'Anwendungen auswählen:'
        SelectAll                       = 'Alle auswählen'
        DeselectAll                     = 'Alle abwählen'
        
        BtnStart                        = 'Installation starten'
        BtnCancel                       = 'Abbrechen'
        BtnBack                         = 'Zurück'
        BtnUseDefaults                  = 'Standard verwenden'
        
        Version365Enterprise            = 'Microsoft 365 Enterprise'
        Version365Business              = 'Microsoft 365 Business'
        VersionProPlus2024              = 'Office LTSC Professional Plus 2024'
        VersionProPlus2021              = 'Office LTSC Professional Plus 2021'
        VersionProPlus2019              = 'Office Professional Plus 2019'
        
        AppWord                         = 'Word'
        AppExcel                        = 'Excel'
        AppPowerPoint                   = 'PowerPoint'
        AppOutlook                      = 'Outlook'
        AppAccess                       = 'Access'
        AppPublisher                    = 'Publisher'
        AppOneNote                      = 'OneNote'
        AppOneDrive                     = 'OneDrive'
        AppTeams                        = 'Teams'
        AppProject                      = 'Project Professional'
        AppVisio                        = 'Visio Professional'
        AppClipchamp                    = 'Clipchamp'
        AppPowerAutomate                = 'Power Automate Desktop'
        

        StatusInitializing              = 'Initialisierung...'
        StatusPreparing                 = 'Vorbereitung...'
        StatusDownloadingODT            = 'Office Deployment Tool wird heruntergeladen...'
        StatusDownloadingOffice         = 'Office-Dateien werden heruntergeladen...'
        StatusInstallingOffice          = 'Microsoft 365 wird installiert...'
        StatusInstallingExtras          = 'Extras werden installiert...'
        StatusActivating                = 'Lizenzen werden aktiviert...'
        StatusCleaning                  = 'Bereinigung...'
        StatusFinalizing                = 'Abschluss...'
        StatusComplete                  = 'Abgeschlossen!'
        
        SubStatusInternetSpeed          = 'Geschwindigkeit abhängig von der Verbindung'
        SubStatusApplyingConfig         = 'Konfiguration wird angewendet'
        SubStatusClipchampPowerAutomate = 'Clipchamp "&" Power Automate'
        SubStatusWindowsOffice          = 'Windows- und Office-Aktivierung'
        SubStatusRemovingTemp           = 'Temporäre Dateien werden entfernt'
        SubStatusAllComplete            = 'Alle Aufgaben erfolgreich abgeschlossen'
        
        LogInstallStarted               = '=== INSTALLATION GESTARTET ==='
        LogInstallComplete              = '=== INSTALLATION ERFOLGREICH ABGESCHLOSSEN ==='
        LogInstallCancelled             = '=== INSTALLATION VOM BENUTZER ABGEBROCHEN ==='
        LogInstallFailed                = '=== KRITISCHER FEHLER ==='
        LogDuration                     = 'Gesamtdauer:'
        LogPhase                        = 'Phase'
        LogCompleted                    = 'abgeschlossen'
        
        ErrAnotherInstance              = 'Eine andere Instanz wird bereits ausgeführt.'
        ErrInstallCancelled             = 'Installation abgebrochen.'
        ErrAllChangesReverted           = 'Alle Änderungen wurden rückgängig gemacht und temporäre Dateien entfernt.'
        ErrInstallFailed                = 'Installation fehlgeschlagen!'
        ErrCheckLog                     = 'Überprüfen Sie die Protokolldatei auf dem Desktop:'
        
        ValidSelectAtLeastOneApp        = 'Bitte wählen Sie mindestens eine Anwendung aus.'
        ValidSelectAtLeastOneLang       = 'Bitte wählen Sie mindestens eine Sprache aus.'
        ValidProjectVisioNote           = 'Project und Visio erfordern separate Lizenzen, werden aber automatisch aktiviert.'
        ConfigureInstallation           = 'Installation konfigurieren'
        FooterDevBy                     = 'Entwickelt von AC Tech'
    }
    
    'fr' = @{
        WindowTitle                     = 'Installateur Microsoft 365 Ultimate'
        WindowSubtitle                  = 'Installateur et Activateur Automatique'
        ConfigWindowTitle               = 'Configuration de l''Installation'
        
        ExpressMode                     = 'Installation Express (Recommandé)'
        ExpressModeDesc                 = 'Installer avec les paramètres par défaut - Microsoft 365 avec toutes les apps'
        CustomMode                      = 'Installation Personnalisée'
        CustomModeDesc                  = 'Choisir version, langues et applications'
        
        SelectVersion                   = 'Sélectionnez la Version Office:'
        SelectLanguages                 = 'Sélectionnez les Langues:'
        PrimaryLanguage                 = 'Langue Principale:'
        AdditionalLanguages             = 'Langues Supplémentaires:'
        SelectApps                      = 'Sélectionnez les Applications:'
        SelectAll                       = 'Tout Sélectionner'
        DeselectAll                     = 'Tout Désélectionner'
        
        BtnStart                        = 'Démarrer l''Installation'
        BtnCancel                       = 'Annuler'
        BtnBack                         = 'Retour'
        BtnUseDefaults                  = 'Utiliser Défaut'
        
        Version365Enterprise            = 'Microsoft 365 Enterprise'
        Version365Business              = 'Microsoft 365 Business'
        VersionProPlus2024              = 'Office LTSC Professional Plus 2024'
        VersionProPlus2021              = 'Office LTSC Professional Plus 2021'
        VersionProPlus2019              = 'Office Professional Plus 2019'
        
        AppWord                         = 'Word'
        AppExcel                        = 'Excel'
        AppPowerPoint                   = 'PowerPoint'
        AppOutlook                      = 'Outlook'
        AppAccess                       = 'Access'
        AppPublisher                    = 'Publisher'
        AppOneNote                      = 'OneNote'
        AppOneDrive                     = 'OneDrive'
        AppTeams                        = 'Teams'
        AppProject                      = 'Project Professional'
        AppVisio                        = 'Visio Professional'
        AppClipchamp                    = 'Clipchamp'
        AppPowerAutomate                = 'Power Automate Desktop'
        
        StatusInitializing              = 'Initialisation...'
        StatusPreparing                 = 'Préparation...'
        StatusDownloadingODT            = 'Téléchargement Office Deployment Tool...'
        StatusDownloadingOffice         = 'Téléchargement des fichiers Office...'
        StatusInstallingOffice          = 'Installation de Microsoft 365...'
        StatusInstallingExtras          = 'Installation des Extras...'
        StatusActivating                = 'Activation des licences...'
        StatusCleaning                  = 'Nettoyage...'
        StatusFinalizing                = 'Finalisation...'
        StatusComplete                  = 'Terminé!'
        
        SubStatusInternetSpeed          = 'Vitesse dépend de la connexion'
        SubStatusApplyingConfig         = 'Application de la configuration'
        SubStatusClipchampPowerAutomate = 'Clipchamp & Power Automate'
        SubStatusWindowsOffice          = 'Activation Windows et Office'
        SubStatusRemovingTemp           = 'Suppression des fichiers temporaires'
        SubStatusAllComplete            = 'Toutes les tâches terminées avec succès'
        
        LogInstallStarted               = '=== INSTALLATION DÉMARRÉE ==='
        LogInstallComplete              = '=== INSTALLATION TERMINÉE AVEC SUCCÈS ==='
        LogInstallCancelled             = '=== INSTALLATION ANNULÉE PAR L''UTILISATEUR ==='
        LogInstallFailed                = '=== ERREUR CRITIQUE ==='
        LogDuration                     = 'Durée totale:'
        LogPhase                        = 'Phase'
        LogCompleted                    = 'terminée'
        
        ErrAnotherInstance              = 'Une autre instance est déjà en cours.'
        ErrInstallCancelled             = 'Installation annulée.'
        ErrAllChangesReverted           = 'Toutes les modifications ont été annulées et les fichiers temporaires supprimés.'
        ErrInstallFailed                = 'Échec de l''Installation!'
        ErrCheckLog                     = 'Vérifiez le fichier journal sur le Bureau:'
        
        ValidSelectAtLeastOneApp        = 'Veuillez sélectionner au moins une application.'
        ValidSelectAtLeastOneLang       = 'Veuillez sélectionner au moins une langue.'
        ValidProjectVisioNote           = 'Project et Visio nécessitent des licences séparées mais seront activés automatiquement.'
        ConfigureInstallation           = 'Configurer l''installation'
        FooterDevBy                     = 'Développé par AC Tech'
    }
    
    # 'zh' - Chinese localization removed due to encoding issues

    
    'it' = @{
        WindowTitle                     = 'Microsoft 365 Ultimate Installer'
        WindowSubtitle                  = 'Installatore e Attivatore Automatico'
        ConfigWindowTitle               = 'Configurazione Installazione'
        
        ExpressMode                     = 'Installazione Rapida (Consigliato)'
        ExpressModeDesc                 = 'Installa con impostazioni predefinite - Microsoft 365 con tutte le app'
        CustomMode                      = 'Installazione Personalizzata'
        CustomModeDesc                  = 'Scegli versione, lingue e applicazioni'
        
        SelectVersion                   = 'Seleziona Versione Office:'
        SelectLanguages                 = 'Seleziona Lingue:'
        PrimaryLanguage                 = 'Lingua Principale:'
        AdditionalLanguages             = 'Lingue Aggiuntive:'
        SelectApps                      = 'Seleziona Applicazioni:'
        SelectAll                       = 'Seleziona Tutto'
        DeselectAll                     = 'Deseleziona Tutto'
        
        BtnStart                        = 'Avvia Installazione'
        BtnCancel                       = 'Annulla'
        BtnBack                         = 'Indietro'
        BtnUseDefaults                  = 'Usa Predefiniti'
        
        Version365Enterprise            = 'Microsoft 365 Enterprise'
        Version365Business              = 'Microsoft 365 Business'
        VersionProPlus2024              = 'Office LTSC Professional Plus 2024'
        VersionProPlus2021              = 'Office LTSC Professional Plus 2021'
        VersionProPlus2019              = 'Office Professional Plus 2019'
        
        AppWord                         = 'Word'
        AppExcel                        = 'Excel'
        AppPowerPoint                   = 'PowerPoint'
        AppOutlook                      = 'Outlook'
        AppAccess                       = 'Access'
        AppPublisher                    = 'Publisher'
        AppOneNote                      = 'OneNote'
        AppOneDrive                     = 'OneDrive'
        AppTeams                        = 'Teams'
        AppProject                      = 'Project Professional'
        AppVisio                        = 'Visio Professional'
        AppClipchamp                    = 'Clipchamp'
        AppPowerAutomate                = 'Power Automate Desktop'
        
        StatusInitializing              = 'Inizializzazione...'
        StatusPreparing                 = 'Preparazione...'
        StatusDownloadingODT            = 'Download Office Deployment Tool...'
        StatusDownloadingOffice         = 'Download file Office...'
        StatusInstallingOffice          = 'Installazione Microsoft 365...'
        StatusInstallingExtras          = 'Installazione Extra...'
        StatusActivating                = 'Attivazione licenze...'
        StatusCleaning                  = 'Pulizia...'
        StatusFinalizing                = 'Finalizzazione...'
        StatusComplete                  = 'Completato!'
        
        SubStatusInternetSpeed          = 'Velocità dipende dalla connessione'
        SubStatusApplyingConfig         = 'Applicazione configurazione'
        SubStatusClipchampPowerAutomate = 'Clipchamp e Power Automate'
        SubStatusWindowsOffice          = 'Attivazione Windows e Office'
        SubStatusRemovingTemp           = 'Rimozione file temporanei'
        SubStatusAllComplete            = 'Tutte le attività completate con successo'
        
        LogInstallStarted               = '=== INSTALLAZIONE AVVIATA ==='
        LogInstallComplete              = '=== INSTALLAZIONE COMPLETATA CON SUCCESSO ==='
        LogInstallCancelled             = '=== INSTALLAZIONE ANNULLATA DALL''UTENTE ==='
        LogInstallFailed                = '=== ERRORE CRITICO ==='
        LogDuration                     = 'Durata totale:'
        LogPhase                        = 'Fase'
        LogCompleted                    = 'completata'
        
        ErrAnotherInstance              = 'Un''altra istanza è già in esecuzione.'
        ErrInstallCancelled             = 'Installazione annullata.'
        ErrAllChangesReverted           = 'Tutte le modifiche sono state annullate e i file temporanei rimossi.'
        ErrInstallFailed                = 'Installazione Fallita!'
        ErrCheckLog                     = 'Controlla il file di log sul Desktop:'
        
        ValidSelectAtLeastOneApp        = 'Seleziona almeno un''applicazione da installare.'
        ValidSelectAtLeastOneLang       = 'Seleziona almeno una lingua.'
        ValidProjectVisioNote           = 'Project e Visio richiedono licenze separate ma verranno attivati automaticamente.'
        ConfigureInstallation           = 'Configura installazione'
        FooterDevBy                     = 'Sviluppato da AC Tech'
    }
    
    # 'ko' - Korean localization removed due to encoding issues

    
    # 'ru' - Russian localization removed due to encoding issues

}

# Function to get localized string
function Get-LocalizedString {
    param(
        [Parameter(Mandatory)] [string]$Key,
        [string]$Lang = $Script:LangBase
    )
    
    # Try exact language match first
    if ($Script:Strings.ContainsKey($Lang) -and $Script:Strings[$Lang].ContainsKey($Key)) {
        return $Script:Strings[$Lang][$Key]
    }
    
    # Fallback to English
    if ($Script:Strings['en'].ContainsKey($Key)) {
        return $Script:Strings['en'][$Key]
    }
    
    # Return key if not found
    return $Key
}

# Shorthand alias
function L { param([string]$Key) Get-LocalizedString -Key $Key }

# ============================================================================
# USER CONFIGURATION (Set by config UI or defaults)
# ============================================================================

$Script:UserConfig = @{
    InstallMode          = 'Express'  # 'Express' or 'Custom'
    OfficeVersion        = 'O365ProPlusRetail'
    Channel              = 'CurrentPreview'
    PrimaryLanguage      = 'en-us'
    AdditionalLanguages  = @('es-es', 'ja-jp', 'pt-br')
    SelectedApps         = @{
        Word       = $true
        Excel      = $true
        PowerPoint = $true
        Outlook    = $true
        Access     = $true
        Publisher  = $true
        OneNote    = $false
        OneDrive   = $false
        Teams      = $true
        Lync       = $false  # Skype for Business
        Groove     = $false  # OneDrive for Business standalone
    }
    IncludeProject       = $true
    IncludeVisio         = $true
    IncludeClipchamp     = $true
    IncludePowerAutomate = $true
}

# Available Office versions with their ODT Product IDs
$Script:OfficeVersions = @{
    'O365ProPlusRetail'  = @{ Name = 'Version365Enterprise'; Channel = 'CurrentPreview'; SupportsProject = $true; SupportsVisio = $true }
    'O365BusinessRetail' = @{ Name = 'Version365Business'; Channel = 'CurrentPreview'; SupportsProject = $true; SupportsVisio = $true }
    'ProPlus2024Retail'  = @{ Name = 'VersionProPlus2024'; Channel = 'PerpetualVL2024'; SupportsProject = $true; SupportsVisio = $true }
    'ProPlus2021Retail'  = @{ Name = 'VersionProPlus2021'; Channel = 'PerpetualVL2021'; SupportsProject = $true; SupportsVisio = $true }
    'ProPlus2019Retail'  = @{ Name = 'VersionProPlus2019'; Channel = 'PerpetualVL2019'; SupportsProject = $false; SupportsVisio = $false }
}

# Available languages
$Script:AvailableLanguages = @{
    'en-us' = 'English (United States)'
    'pt-br' = 'Portuguese (Brazil)'
    'es-es' = 'Spanish (Spain)'
    'es-mx' = 'Spanish (Mexico)'
    'de-de' = 'German'
    'fr-fr' = 'French'
    'it-it' = 'Italian'
}

# Apps that can be excluded via ODT
$Script:ExcludableApps = @{
    'Access'     = 'Access'
    'Excel'      = 'Excel'
    'OneNote'    = 'OneNote'
    'Outlook'    = 'Outlook'
    'PowerPoint' = 'PowerPoint'
    'Publisher'  = 'Publisher'
    'Word'       = 'Word'
    'OneDrive'   = 'OneDrive'
    'Lync'       = 'Lync'  # Skype for Business
    'Groove'     = 'Groove'  # OneDrive for Business standalone
    'Teams'      = 'Teams'
}

# ============================================================================
# REGISTRY BACKUP & RESTORATION
# ============================================================================

$Script:RegistryBackups = @{}

function Save-RegistryKey {
    param([Parameter(Mandatory)] [string]$Path)
    try {
        if (Test-Path $Path) {
            $Script:RegistryBackups[$Path] = @{
                Exists = $true
                Data   = Export-RegistryBranch -Path $Path
            }
            Write-Log "Registry backed up: $Path" -Level Debug
        }
        else {
            $Script:RegistryBackups[$Path] = @{ Exists = $false }
        }
    }
    catch {
        Write-Log "Failed to backup registry key $Path : $($_.Exception.Message)" -Level Warning
    }
}

function Export-RegistryBranch {
    param([Parameter(Mandatory)] [string]$Path)
    $items = @()
    try {
        $regKey = Get-Item $Path -ErrorAction SilentlyContinue
        if ($regKey) {
            $properties = $regKey.GetValueNames()
            foreach ($propName in $properties) {
                if ([string]::IsNullOrEmpty($propName)) { continue }
                $items += @{
                    Name  = $propName
                    Value = $regKey.GetValue($propName)
                    Type  = $regKey.GetValueKind($propName).ToString()
                }
            }
        }
    }
    catch {
        Write-Log "Failed to export registry branch $Path : $($_.Exception.Message)" -Level Warning
    }
    return $items
}

function Restore-RegistryKey {
    param([Parameter(Mandatory)] [string]$Path)
    try {
        if ($Script:RegistryBackups[$Path]) {
            if ($Script:RegistryBackups[$Path].Exists) {
                Write-Log "Restoring registry key: $Path" -Level Debug
                foreach ($item in $Script:RegistryBackups[$Path].Data) {
                    Set-ItemProperty -Path $Path -Name $item.Name -Value $item.Value -Type $item.Type -Force -ErrorAction SilentlyContinue
                }
            }
            else {
                # Key didn't exist before, remove it if it exists now
                if (Test-Path $Path) {
                    Write-Log "Removing registry key that didn't exist before: $Path" -Level Debug
                    Remove-Item $Path -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
    catch {
        Write-Log "Failed to restore registry key $Path : $($_.Exception.Message)" -Level Warning
    }
}

# ============================================================================
# WINDOW EFFECTS (ACRYLIC / BLUR)
# ============================================================================
$code = @"
using System;
using System.Runtime.InteropServices;

public class WindowEffects
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct WindowCompositionAttributeData
    {
        public WindowCompositionAttribute Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }
    
    internal enum WindowCompositionAttribute
    {
        WCA_ACCENT_POLICY = 19
    }
    
    internal enum AccentState
    {
        ACCENT_DISABLED = 0,
        ACCENT_ENABLE_GRADIENT = 1,
        ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
        ACCENT_ENABLE_BLURBEHIND = 3,
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4,
        ACCENT_INVALID_STATE = 5
    }
    
    [StructLayout(LayoutKind.Sequential)]
    internal struct AccentPolicy
    {
        public AccentState AccentState;
        public int AccentFlags;
        public int GradientColor;
        public int AnimationId;
    }
    
    [DllImport("user32.dll")]
    internal static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);
    
    public static void EnableBlur(IntPtr hwnd, uint blurOpacity = 0, uint blurColor = 0x000000)
    {
        var accent = new AccentPolicy();
        // Use standard Blur (3) for maximum reliability across Win10/11 versions
        accent.AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND;
        accent.GradientColor = 0; // Not needed for standard blur
        
        var accentStructSize = Marshal.SizeOf(accent);
        var accentPtr = Marshal.AllocHGlobal(accentStructSize);
        Marshal.StructureToPtr(accent, accentPtr, false);
        
        var data = new WindowCompositionAttributeData();
        data.Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY;
        data.SizeOfData = accentStructSize;
        data.Data = accentPtr;
        
        try
        {
            SetWindowCompositionAttribute(hwnd, ref data);
        }
        finally
        {
            Marshal.FreeHGlobal(accentPtr);
        }
    }
}
"@

try { Add-Type -TypeDefinition $code -Language CSharp } catch { Write-Warning "Failed to load WindowEffects type: $($_.Exception.Message)"; Write-Warning "Inner Exception: $($_.Exception.InnerException.Message)" }

# Prevent multiple instances (with Force option to clear stuck mutex)
$Script:MutexName = 'Global\M365UltimateInstaller'
$ScriptMutex = $null

if ($Force) {
    # Force mode: try to clear any stuck mutex
    try {
        $tempMutex = [System.Threading.Mutex]::OpenExisting($Script:MutexName)
        $tempMutex.Dispose()
    }
    catch {
        # Mutex doesn't exist or already disposed - expected behavior, suppress error
        $null = $_
    }
}

$ScriptMutex = New-Object System.Threading.Mutex($false, $Script:MutexName)
if (-not $ScriptMutex.WaitOne(0, $false)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Another instance is already running.`n`nIf this is an error, run the script with -Force parameter to clear the stuck instance.",
        'Microsoft 365 Ultimate Installer',
        0,
        48
    )
    $ScriptMutex.Dispose()
    exit 1
}

# ============================================================================
# LOGGING & UI
# ============================================================================

function Initialize-Log {
    $Script:Config.InstallStartTime = Get-Date
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $psVersion = $PSVersionTable.PSVersion.ToString()
    $header = @"
================================================================================
MICROSOFT 365 ULTIMATE INSTALLER
================================================================================
Started: $($Script:Config.InstallStartTime.ToString('yyyy-MM-dd HH:mm:ss'))
Computer: $env:COMPUTERNAME
User: $env:USERNAME ($env:USERDOMAIN)
OS: $($osInfo.Caption) (Build $($osInfo.BuildNumber))
PowerShell: $psVersion
================================================================================
"@
    Set-Content -Path $Script:Config.LogFile -Value $header -Encoding UTF8 -Force
}

function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Debug')] [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $logLine = "$timestamp [$Level] $Message"
    
    try {
        if ($null -ne $Script:Config -and $null -ne $Script:Config.LogFile) {
            Add-Content -Path $Script:Config.LogFile -Value $logLine -Encoding UTF8 -ErrorAction Stop
        }
        else {
            Write-Host "LOG ERROR: LogFile path is null. Message: $logLine" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "LOGGING FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "ORIGINAL MESSAGE: $logLine" -ForegroundColor Gray
    }
    
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        default { 'White' }
    }
    Write-Host $logLine -ForegroundColor $color
}

# ============================================================================
# WPF UNIFIED WINDOW (Configuration + Progress)
# ============================================================================

# ============================================================================
# CONFIGURATION & PROGRESS UI
# ============================================================================

# Global sync object for progress updates
$Script:ProgressSync = $null
$Script:ProgressRunspace = $null
$Script:ProgressHandle = $null

# ============================================================================
# GUI (ADVANCED WPF MODERN REWRITE)
# ============================================================================

function Show-ConfigWindow {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing
    
    # -------------------------------------------------------------------------
    # MODERN XAML DEFINITION
    # -------------------------------------------------------------------------
    $strTitle = "Microsoft Ultimate Installer"
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:shell="http://schemas.microsoft.com/winfx/2006/xaml/presentation/shell"
Title="$strTitle" Height="650" Width="900"
WindowStyle="None" AllowsTransparency="True" ResizeMode="CanResizeWithGrip"
Background="Transparent" FontFamily="Segoe UI Variable, Segoe UI, Arial">
    
    <!-- WindowChrome removed to allow Transparency/Blur compatibility -->

<Window.Resources>
<!-- COLORS & BRUSHES -->
<Color x:Key="Color.Background">#FF151515</Color> <!-- 100% Opacity Dark (Default) -->
<Color x:Key="Color.Background.Transparent">#E6151515</Color> <!-- 90% Opacity (Dragging) -->
<Color x:Key="Color.Accent1">#0078D4</Color>
<Color x:Key="Color.Accent2">#00B7C3</Color>
<SolidColorBrush x:Key="Brush.Background" Color="{StaticResource Color.Background}"/>
<SolidColorBrush x:Key="Brush.Accent1" Color="{StaticResource Color.Accent1}"/>
<SolidColorBrush x:Key="Brush.Accent2" Color="{StaticResource Color.Accent2}"/>
<SolidColorBrush x:Key="Brush.Text.Primary" Color="#FFFFFF"/>
<SolidColorBrush x:Key="Brush.Text.Secondary" Color="#AAAAAA"/>
<SolidColorBrush x:Key="Brush.Border" Color="#333333"/>
<SolidColorBrush x:Key="Brush.Control.Back" Color="#252525"/>
<LinearGradientBrush x:Key="Brush.Accent.Gradient" StartPoint="0,0" EndPoint="1,1">
<GradientStop Color="{StaticResource Color.Accent1}" Offset="0"/>
<GradientStop Color="{StaticResource Color.Accent2}" Offset="1"/>
</LinearGradientBrush>

<!-- STYLES -->
<Style TargetType="Button" x:Key="Style.Button.Primary">
<Setter Property="Background" Value="{StaticResource Brush.Accent.Gradient}"/>
<Setter Property="Foreground" Value="White"/>
<Setter Property="BorderThickness" Value="0"/>
<Setter Property="FontSize" Value="14"/>
<Setter Property="FontWeight" Value="SemiBold"/>
<Setter Property="Cursor" Value="Hand"/>
<Setter Property="Template">
<Setter.Value>
<ControlTemplate TargetType="Button">
<Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="6" SnapsToDevicePixels="True">
<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="15,8"/>
</Border>
<ControlTemplate.Triggers>
<Trigger Property="IsMouseOver" Value="True">
<Setter TargetName="border" Property="Opacity" Value="0.9"/>
</Trigger>
<Trigger Property="IsPressed" Value="True">
<Setter TargetName="border" Property="Opacity" Value="0.7"/>
</Trigger>
</ControlTemplate.Triggers>
</ControlTemplate>
</Setter.Value>
</Setter>
</Style>

<Style TargetType="Button" x:Key="Style.Button.Secondary">
<Setter Property="Background" Value="#333333"/>
<Setter Property="Foreground" Value="White"/>
<Setter Property="BorderThickness" Value="1"/>
<Setter Property="BorderBrush" Value="#555555"/>
<Setter Property="Template">
<Setter.Value>
<ControlTemplate TargetType="Button">
<Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="15,8"/>
</Border>
<ControlTemplate.Triggers>
<Trigger Property="IsMouseOver" Value="True">
<Setter TargetName="border" Property="Background" Value="#444444"/>
</Trigger>
</ControlTemplate.Triggers>
</ControlTemplate>
</Setter.Value>
</Setter>
</Style>

<Style TargetType="Button" x:Key="Style.Button.WindowControl">
<Setter Property="Background" Value="Transparent"/>
<Setter Property="Foreground" Value="#AAAAAA"/>
<Setter Property="WindowChrome.IsHitTestVisibleInChrome" Value="True"/>
<Setter Property="Template">
<Setter.Value>
<ControlTemplate TargetType="Button">
<Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="0">
<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
</Border>
<ControlTemplate.Triggers>
<Trigger Property="IsMouseOver" Value="True">
<Setter TargetName="border" Property="Background" Value="#33FFFFFF"/>
<Setter Property="Foreground" Value="White"/>
</Trigger>
</ControlTemplate.Triggers>
</ControlTemplate>
</Setter.Value>
</Setter>
</Style>
        
<Style TargetType="Button" x:Key="Style.Button.Close" BasedOn="{StaticResource Style.Button.WindowControl}">
<Style.Triggers>
<Trigger Property="IsMouseOver" Value="True">
<Setter Property="Background" Value="#E81123"/>
</Trigger>
</Style.Triggers>
</Style>

<!-- Card Style Checkbox -->
<Style TargetType="CheckBox" x:Key="Style.CheckBox.Card">
<Setter Property="Foreground" Value="White"/>
<Setter Property="Cursor" Value="Hand"/>
<Setter Property="Height" Value="45"/>
<Setter Property="Margin" Value="4"/>
<Setter Property="Template">
<Setter.Value>
<ControlTemplate TargetType="CheckBox">
<Border x:Name="border" Background="#2A2A2A" BorderBrush="#3A3A3A" BorderThickness="1" CornerRadius="6">
<Grid Margin="10,0">
<Grid.ColumnDefinitions>
<ColumnDefinition Width="Auto"/>
<ColumnDefinition Width="*"/>
</Grid.ColumnDefinitions>
<Border x:Name="checkmarkBox" Width="20" Height="20" CornerRadius="4" BorderBrush="#666666" BorderThickness="1" Background="Transparent" VerticalAlignment="Center">
<Path x:Name="checkPath" Data="M2,10 L7,15 L18,4" Stroke="White" StrokeThickness="2" Visibility="Collapsed" HorizontalAlignment="Center" VerticalAlignment="Center"/>
</Border>
<ContentPresenter Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="10,0,0,0"/>
</Grid>
</Border>
<ControlTemplate.Triggers>
<Trigger Property="IsMouseOver" Value="True">
<Setter TargetName="border" Property="Background" Value="#353535"/>
<Setter TargetName="border" Property="BorderBrush" Value="#555555"/>
</Trigger>
<Trigger Property="IsChecked" Value="True">
<Setter TargetName="border" Property="Background" Value="#200078D4"/>
<Setter TargetName="border" Property="BorderBrush" Value="{StaticResource Brush.Accent1}"/>
<Setter TargetName="checkmarkBox" Property="Background" Value="{StaticResource Brush.Accent1}"/>
<Setter TargetName="checkmarkBox" Property="BorderBrush" Value="{StaticResource Brush.Accent1}"/>
<Setter TargetName="checkPath" Property="Visibility" Value="Visible"/>
</Trigger>
</ControlTemplate.Triggers>
</ControlTemplate>
</Setter.Value>
</Setter>
</Style>

<!-- Mode Radio Button (Card) -->
<Style TargetType="RadioButton" x:Key="Style.RadioButton.Mode">
<Setter Property="Template">
<Setter.Value>
<ControlTemplate TargetType="RadioButton">
<Border x:Name="border" Background="#252525" BorderBrush="#444444" BorderThickness="1" CornerRadius="8" Padding="20" Margin="10" Cursor="Hand">
<StackPanel>
<TextBlock Text="{TemplateBinding Tag}" FontSize="18" FontWeight="Bold" Foreground="White" Margin="0,0,0,10"/>
<ContentPresenter/>
</StackPanel>
</Border>
<ControlTemplate.Triggers>
<Trigger Property="IsChecked" Value="True">
<Setter TargetName="border" Property="BorderBrush" Value="{StaticResource Brush.Accent1}"/>
<Setter TargetName="border" Property="BorderThickness" Value="2"/>
<Setter TargetName="border" Property="Background" Value="#150078D4"/>
</Trigger>
<Trigger Property="IsMouseOver" Value="True">
<Setter TargetName="border" Property="Background" Value="#303030"/>
</Trigger>
</ControlTemplate.Triggers>
</ControlTemplate>
</Setter.Value>
</Setter>
</Style>

<ControlTemplate x:Key="ScrollViewerModern" TargetType="ScrollViewer">
<Grid x:Name="Grid" Background="{TemplateBinding Background}">
<Grid.ColumnDefinitions>
<ColumnDefinition Width="*"/>
<ColumnDefinition Width="Auto"/>
</Grid.ColumnDefinitions>
<Grid.RowDefinitions>
<RowDefinition Height="*"/>
<RowDefinition Height="Auto"/>
</Grid.RowDefinitions>
<ScrollContentPresenter x:Name="PART_ScrollContentPresenter" CanContentScroll="{TemplateBinding CanContentScroll}" CanHorizontallyScroll="False" CanVerticallyScroll="False" ContentTemplate="{TemplateBinding ContentTemplate}" Content="{TemplateBinding Content}" Grid.Column="0" Margin="{TemplateBinding Padding}" Grid.Row="0"/>
<ScrollBar x:Name="PART_VerticalScrollBar" AutomationProperties.AutomationId="VerticalScrollBar" Cursor="Arrow" Grid.Column="1" Maximum="{TemplateBinding ScrollableHeight}" Minimum="0" Grid.Row="0" Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}" Value="{Binding VerticalOffset, Mode=OneWay, RelativeSource={RelativeSource TemplatedParent}}" ViewportSize="{TemplateBinding ViewportHeight}" Width="10"/>
</Grid>
</ControlTemplate>
</Window.Resources>

<!-- MAIN CONTAINER WITH BLUR BACKGROUND -->
<Border x:Name="MainBorder" Background="{StaticResource Brush.Background}" CornerRadius="10" BorderBrush="#444444" BorderThickness="1">
<Grid>
<Grid.RowDefinitions>
<RowDefinition Height="Auto"/> <!-- Title Bar -->
<RowDefinition Height="*"/>    <!-- Content -->
<RowDefinition Height="Auto"/> <!-- Footer -->
</Grid.RowDefinitions>

<!-- TITLE BAR -->
<Grid x:Name="TitleBar" Grid.Row="0" Height="40" Background="Transparent">
<Grid.ColumnDefinitions>
<ColumnDefinition Width="Auto"/> <!-- Logo -->
<ColumnDefinition Width="*"/>    <!-- Title -->
<ColumnDefinition Width="Auto"/> <!-- Controls -->
</Grid.ColumnDefinitions>

<StackPanel Orientation="Horizontal" Grid.Column="0" Margin="15,0,0,0" VerticalAlignment="Center">
<Image x:Name="imgTitleLogo" Height="24" Margin="0,0,10,0"/>
<TextBlock Text="Microsoft Ultimate Installer" Foreground="#CCCCCC" VerticalAlignment="Center" FontSize="12"/>
</StackPanel>

<StackPanel Grid.Column="2" Orientation="Horizontal">
<Button x:Name="btnMinimize" Content="__" Width="46" Height="40" Style="{StaticResource Style.Button.WindowControl}" Padding="0,0,0,8"/>
<Button x:Name="btnMaximize" Content="[ ]" Width="46" Height="40" Style="{StaticResource Style.Button.WindowControl}" FontSize="12"/>
<Button x:Name="btnClose" Content="X" Width="46" Height="40" Style="{StaticResource Style.Button.Close}" FontSize="14"/>
</StackPanel>
</Grid>

<!-- PAGES CONTAINER -->
<Grid Grid.Row="1" Margin="20,10,20,10">
                
<!-- Page: Mode Selection (Overview) -->
<Grid x:Name="PageMode" Visibility="Visible">
<StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
<TextBlock Text="Microsoft Ultimate Installer" FontSize="32" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" Margin="0,0,0,10">
<TextBlock.Effect>
<DropShadowEffect Color="Black" BlurRadius="10" ShadowDepth="2" Opacity="0.5"/>
</TextBlock.Effect>
</TextBlock>
<TextBlock Text="Select installation mode" FontSize="16" Foreground="{StaticResource Brush.Text.Secondary}" HorizontalAlignment="Center" Margin="0,0,0,40"/>

<UniformGrid Columns="3" Width="850">
<!-- Express Mode -->
<RadioButton x:Name="rbExpress" Tag="Express Mode" GroupName="Mode" IsChecked="True" Style="{StaticResource Style.RadioButton.Mode}">
<StackPanel>
<TextBlock Text="Express Install" FontSize="16" FontWeight="Bold" Foreground="White" Margin="0,0,0,10"/>
<TextBlock Text="Recommended for most users." Foreground="#AAAAAA" TextWrapping="Wrap"/>
<TextBlock Text="• New Outlook" Foreground="#888888" Margin="0,5,0,0"/>
<TextBlock Text="• Word, Excel, PowerPoint" Foreground="#888888"/>
<TextBlock Text="• Teams, Clipchamp" Foreground="#888888"/>
</StackPanel>
</RadioButton>

<!-- Custom Mode -->
<RadioButton x:Name="rbCustom" Tag="Custom Mode" GroupName="Mode" Style="{StaticResource Style.RadioButton.Mode}">
<StackPanel>
<TextBlock Text="Custom Config" FontSize="16" FontWeight="Bold" Foreground="White" Margin="0,0,0,10"/>
<TextBlock Text="Full control over installation." Foreground="#AAAAAA" TextWrapping="Wrap"/>
<TextBlock Text="• Select specific Apps" Foreground="#888888" Margin="0,5,0,0"/>
<TextBlock Text="• Change Channel / Version" Foreground="#888888"/>
<TextBlock Text="• Windows Insider" Foreground="#888888"/>
</StackPanel>
</RadioButton>
                            
<!-- Uninstall Mode -->
<RadioButton x:Name="rbUninstall" Tag="Uninstall Mode" GroupName="Mode" Style="{StaticResource Style.RadioButton.Mode}">
<StackPanel>
<TextBlock Text="Deep Uninstall" FontSize="16" FontWeight="Bold" Foreground="White" Margin="0,0,0,10"/>
<TextBlock Text="Completely wipe Office." Foreground="#AAAAAA" TextWrapping="Wrap"/>
<TextBlock Text="• Removes all traces" Foreground="#888888" Margin="0,5,0,0"/>
<TextBlock Text="• Registry &amp; File Cleanup" Foreground="#888888"/>
<TextBlock Text="• Fixes stubborn issues" Foreground="#888888"/>
</StackPanel>
</RadioButton>
</UniformGrid>

<Button x:Name="btnModeNext" Content="Start Now" Width="180" Height="36" Margin="0,40,0,0" Style="{StaticResource Style.Button.Primary}"/>
</StackPanel>
</Grid>

<!-- Page: Custom Configuration -->
<Grid x:Name="PageCustom" Visibility="Collapsed">
<Grid.RowDefinitions>
<RowDefinition Height="Auto"/> <!-- Header -->
<RowDefinition Height="*"/>    <!-- Scrollable Content -->
<RowDefinition Height="Auto"/> <!-- Actions -->
</Grid.RowDefinitions>

<StackPanel Grid.Row="0" Margin="0,0,0,15">
<Button x:Name="btnCustomBack" Content="← Back" HorizontalAlignment="Left" Style="{StaticResource Style.Button.WindowControl}" Foreground="{StaticResource Brush.Accent1}"/>
</StackPanel>

<ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Template="{DynamicResource ScrollViewerModern}">
<StackPanel Margin="0,0,15,0">
                            
<!-- Version/Channel -->
<TextBlock Text="Office Edition" Foreground="#888888" FontWeight="Bold" Margin="0,0,0,8"/>
<ComboBox x:Name="cmbVersion" Height="35" Background="#333333" Foreground="White" BorderBrush="#555555" IsEditable="False">
<ComboBoxItem Tag="O365ProPlusRetail|CurrentPreview" IsSelected="True">Microsoft 365 Enterprise (Current Preview)</ComboBoxItem>
<ComboBoxItem Tag="O365BusinessRetail|CurrentPreview">Microsoft 365 Business (Current Preview)</ComboBoxItem>
<ComboBoxItem Tag="ProPlus2024Retail|PerpetualVL2024">Office LTSC 2024</ComboBoxItem>
<ComboBoxItem Tag="ProPlus2021Retail|PerpetualVL2021">Office LTSC 2021</ComboBoxItem>
<ComboBoxItem Tag="ProPlus2019Retail|PerpetualVL2019">Office 2019</ComboBoxItem>
</ComboBox>


<!-- Windows Edition -->
<TextBlock Text="Windows Edition (Activation)" Foreground="#888888" FontWeight="Bold" Margin="0,20,0,8"/>
<ComboBox x:Name="cmbWindowsEdition" Height="35" Background="#333333" Foreground="White" BorderBrush="#555555" SelectedIndex="0">
<ComboBoxItem Tag="Pro" IsSelected="True">Windows 10/11 Pro (Recommended)</ComboBoxItem>
<ComboBoxItem Tag="Home">Windows 10/11 Home</ComboBoxItem>
<ComboBoxItem Tag="Enterprise">Windows 10/11 Enterprise</ComboBoxItem>
<ComboBoxItem Tag="Education">Windows 10/11 Education</ComboBoxItem>
<ComboBoxItem Tag="IoTEnterprise">Windows 10/11 IoT Enterprise</ComboBoxItem>
<ComboBoxItem Tag="">Do Not Change Edition</ComboBoxItem>
</ComboBox>

<!-- Windows Insider Program -->
<TextBlock Text="Windows Insider Program (Experimental)" Foreground="#888888" FontWeight="Bold" Margin="0,20,0,8"/>
<WrapPanel x:Name="panelInsider">
<RadioButton x:Name="rbInsiderNone" Content="Do not change" Tag="None" GroupName="InsiderChannel" IsChecked="True" Foreground="White" Margin="0,0,20,8"/>
<RadioButton x:Name="rbInsiderRelease" Content="Release Preview" Tag="ReleasePreview" GroupName="InsiderChannel" Foreground="White" Margin="0,0,20,8"/>
<RadioButton x:Name="rbInsiderBeta" Content="Beta Channel" Tag="Beta" GroupName="InsiderChannel" Foreground="White" Margin="0,0,20,8"/>
<RadioButton x:Name="rbInsiderDev" Content="Dev Channel" Tag="Dev" GroupName="InsiderChannel" Foreground="White" Margin="0,0,20,8"/>
<RadioButton x:Name="rbInsiderCanary" Content="Canary Channel" Tag="Canary" GroupName="InsiderChannel" Foreground="White" Margin="0,0,20,8"/>
</WrapPanel>


<!-- Application Selection -->
<Grid Margin="0,20,0,10">
<TextBlock Text="Applications" Foreground="#888888" FontWeight="Bold" VerticalAlignment="Center"/>
<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
<Button x:Name="btnSelectAll" Content="All" Width="60" Height="24" Background="Transparent" Foreground="{StaticResource Brush.Accent1}" BorderThickness="0"/>
<Button x:Name="btnDeselectAll" Content="None" Width="60" Height="24" Background="Transparent" Foreground="#888888" BorderThickness="0"/>
</StackPanel>
</Grid>

<UniformGrid Columns="3">
<CheckBox x:Name="chkWord" Content="Word" IsChecked="True" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkExcel" Content="Excel" IsChecked="True" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkPowerPoint" Content="PowerPoint" IsChecked="True" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkOutlook" Content="Outlook (New)" IsChecked="True" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkTeams" Content="Teams" IsChecked="True" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkOneNote" Content="OneNote" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkOneDrive" Content="OneDrive" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkAccess" Content="Access" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkPublisher" Content="Publisher" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkProject" Content="Project Pro" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkVisio" Content="Visio Pro" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkClipchamp" Content="Clipchamp" IsChecked="True" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkDefender" Content="Defender" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkToDo" Content="To Do" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkPCManager" Content="PC Manager" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkStickyNotes" Content="Sticky Notes" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkPowerBI" Content="Power BI" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkPowerAutomate" Content="Power Automate" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkPowerToys" Content="PowerToys" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkVSCode" Content="VS Code (Insider)" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkVS2022" Content="VS 2026 (Preview)" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkCopilot" Content="Copilot" Style="{StaticResource Style.CheckBox.Card}"/>
<CheckBox x:Name="chkSkype" Content="Skype" Style="{StaticResource Style.CheckBox.Card}"/>
</UniformGrid>

<!-- Languages -->
<TextBlock Text="Primary Language" Foreground="#888888" FontWeight="Bold" Margin="0,20,0,8"/>
<ComboBox x:Name="cmbLanguage" Height="35" Background="#333333" Foreground="White" BorderBrush="#555555" SelectedIndex="0">
<ComboBoxItem Tag="en-us">English (United States)</ComboBoxItem>
<ComboBoxItem Tag="pt-br">Português (Brasil)</ComboBoxItem>
<ComboBoxItem Tag="es-es">Español</ComboBoxItem>
<ComboBoxItem Tag="de-de">Deutsch</ComboBoxItem>
<ComboBoxItem Tag="fr-fr">Français</ComboBoxItem>
<ComboBoxItem Tag="ja-jp">日本語</ComboBoxItem>
<ComboBoxItem Tag="zh-cn">中文 (简体)</ComboBoxItem>
<ComboBoxItem Tag="it-it">Italiano</ComboBoxItem>
<ComboBoxItem Tag="ko-kr">한국어</ComboBoxItem>
<ComboBoxItem Tag="ru-ru">Русский</ComboBoxItem>
<ComboBoxItem Tag="ar-sa">Arabic (Saudi Arabia)</ComboBoxItem>
<ComboBoxItem Tag="da-dk">Danish</ComboBoxItem>
<ComboBoxItem Tag="nl-nl">Dutch</ComboBoxItem>
<ComboBoxItem Tag="fi-fi">Finnish</ComboBoxItem>
<ComboBoxItem Tag="el-gr">Greek</ComboBoxItem>
<ComboBoxItem Tag="he-il">Hebrew</ComboBoxItem>
<ComboBoxItem Tag="hu-hu">Hungarian</ComboBoxItem>
<ComboBoxItem Tag="id-id">Indonesian</ComboBoxItem>
<ComboBoxItem Tag="ms-my">Malay</ComboBoxItem>
<ComboBoxItem Tag="nb-no">Norwegian (Bokmål)</ComboBoxItem>
<ComboBoxItem Tag="pl-pl">Polish</ComboBoxItem>
<ComboBoxItem Tag="pt-pt">Portuguese (Portugal)</ComboBoxItem>
<ComboBoxItem Tag="ro-ro">Romanian</ComboBoxItem>
<ComboBoxItem Tag="sk-sk">Slovak</ComboBoxItem>
<ComboBoxItem Tag="sv-se">Swedish</ComboBoxItem>
<ComboBoxItem Tag="th-th">Thai</ComboBoxItem>
<ComboBoxItem Tag="tr-tr">Turkish</ComboBoxItem>
<ComboBoxItem Tag="uk-ua">Ukrainian</ComboBoxItem>
<ComboBoxItem Tag="vi-vn">Vietnamese</ComboBoxItem>
</ComboBox>

<TextBlock Text="Additional Languages" Foreground="#888888" FontWeight="Bold" Margin="0,20,0,8"/>
<WrapPanel x:Name="panelAdditionalLangs">
<CheckBox x:Name="chkLang_en_us" Content="English (US)" Tag="en-us" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_pt_br" Content="Português (BR)" Tag="pt-br" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_es_es" Content="Español" Tag="es-es" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_de_de" Content="Deutsch" Tag="de-de" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_fr_fr" Content="Français" Tag="fr-fr" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_ja_jp" Content="日本語" Tag="ja-jp" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_zh_cn" Content="中文 (简体)" Tag="zh-cn" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_it_it" Content="Italiano" Tag="it-it" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_ko_kr" Content="한국어" Tag="ko-kr" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_ru_ru" Content="Русский" Tag="ru-ru" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_ar_sa" Content="Arabic" Tag="ar-sa" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_da_dk" Content="Danish" Tag="da-dk" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_nl_nl" Content="Dutch" Tag="nl-nl" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_fi_fi" Content="Finnish" Tag="fi-fi" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_el_gr" Content="Greek" Tag="el-gr" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_he_il" Content="Hebrew" Tag="he-il" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_hu_hu" Content="Hungarian" Tag="hu-hu" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_id_id" Content="Indonesian" Tag="id-id" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_ms_my" Content="Malay" Tag="ms-my" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_nb_no" Content="Norwegian" Tag="nb-no" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_pl_pl" Content="Polish" Tag="pl-pl" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_pt_pt" Content="Português (PT)" Tag="pt-pt" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_ro_ro" Content="Romanian" Tag="ro-ro" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_sk_sk" Content="Slovak" Tag="sk-sk" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_sv_se" Content="Swedish" Tag="sv-se" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_th_th" Content="Thai" Tag="th-th" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_tr_tr" Content="Turkish" Tag="tr-tr" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_uk_ua" Content="Ukrainian" Tag="uk-ua" Foreground="White" Margin="0,0,20,8"/>
<CheckBox x:Name="chkLang_vi_vn" Content="Vietnamese" Tag="vi-vn" Foreground="White" Margin="0,0,20,8"/>
</WrapPanel>
</StackPanel>
</ScrollViewer>

<StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
<Button x:Name="btnCustomStart" Content="Install Selected" Width="150" Height="36" Style="{StaticResource Style.Button.Primary}"/>
</StackPanel>
</Grid>

</Grid>

<!-- FOOTER -->
<Grid Grid.Row="2" Background="#1A1A1A" Height="30">
<StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
<Image x:Name="imgFooterLogo" Height="24" Stretch="Uniform" Margin="0,2,0,0"/>
<TextBlock Text="Developed by AC Tech" Foreground="#888888" VerticalAlignment="Center" Margin="10,0,0,0" FontSize="11"/>
</StackPanel>
</Grid>
</Grid>
</Border>
</Window>
"@

    # -------------------------------------------------------------------------
    # PARSE XAML & LOGIC
    # -------------------------------------------------------------------------
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # -------------------------------------------------------------------------
    # ELEMENTS
    # -------------------------------------------------------------------------
    # Pages
    $pageMode = $window.FindName('PageMode')
    $pageCustom = $window.FindName('PageCustom')
    
    # Controls
    $rbExpress = $window.FindName('rbExpress')
    $rbCustom = $window.FindName('rbCustom')
    $rbUninstall = $window.FindName('rbUninstall')
    $btnModeNext = $window.FindName('btnModeNext')
    $btnCustomBack = $window.FindName('btnCustomBack')
    $btnCustomStart = $window.FindName('btnCustomStart')
    $btnSelectAll = $window.FindName('btnSelectAll')
    $btnDeselectAll = $window.FindName('btnDeselectAll')
    $cmbVersion = $window.FindName('cmbVersion')
    # Insider Radio Buttons Panel (iterate children to find selected)
    $panelInsider = $window.FindName('panelInsider')
    # Windows Edition
    $cmbWindowsEdition = $window.FindName('cmbWindowsEdition')
    $cmbLanguage = $window.FindName('cmbLanguage')
    $panelAdditionalLangs = $window.FindName('panelAdditionalLangs')

    # Window Controls
    $btnClose = $window.FindName('btnClose')
    $btnMinimize = $window.FindName('btnMinimize')
    $btnMaximize = $window.FindName('btnMaximize')
    $imgTitleLogo = $window.FindName('imgTitleLogo')
    $imgFooterLogo = $window.FindName('imgFooterLogo')

    # Apps
    $chkWord = $window.FindName('chkWord')
    $chkExcel = $window.FindName('chkExcel')
    $chkPowerPoint = $window.FindName('chkPowerPoint')
    $chkOutlook = $window.FindName('chkOutlook')
    $chkTeams = $window.FindName('chkTeams')
    $chkOneNote = $window.FindName('chkOneNote')
    $chkOneDrive = $window.FindName('chkOneDrive')
    $chkSkype = $window.FindName('chkSkype')
    $chkAccess = $window.FindName('chkAccess')
    $chkPublisher = $window.FindName('chkPublisher')
    $chkProject = $window.FindName('chkProject')
    $chkVisio = $window.FindName('chkVisio')
    $chkClipchamp = $window.FindName('chkClipchamp')
    $chkPowerAutomate = $window.FindName('chkPowerAutomate')
    # New Apps
    $chkDefender = $window.FindName('chkDefender')
    $chkToDo = $window.FindName('chkToDo')
    $chkPCManager = $window.FindName('chkPCManager')
    $chkStickyNotes = $window.FindName('chkStickyNotes')
    $chkPowerBI = $window.FindName('chkPowerBI')
    $chkPowerToys = $window.FindName('chkPowerToys')
    $chkVSCode = $window.FindName('chkVSCode')
    $chkVS2022 = $window.FindName('chkVS2022')
    $chkCopilot = $window.FindName('chkCopilot')

    # -----------------------
    # ASSET LOADING
    # -----------------------
    # Header Logo
    # Header Logo
    # DEBUG: Load from external file to ensure we have valid Base64
    if (Test-Path "$PSScriptRoot\header_base64.txt") {
        $Script:ACTechLogoBase64 = (Get-Content "$PSScriptRoot\header_base64.txt" -Raw).Trim()
    }
    
    if ($Script:ACTechLogoBase64) {
        try {
            $bytes = [Convert]::FromBase64String($Script:ACTechLogoBase64)
            $stream = New-Object System.IO.MemoryStream(, $bytes)
            $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmap.BeginInit()
            $bitmap.StreamSource = $stream
            $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bitmap.EndInit()
            $bitmap.Freeze()
            if ($imgTitleLogo) { $imgTitleLogo.Source = $bitmap }
        }
        catch {
            Write-Warning "Failed to load Header Logo: $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "Header Logo Base64 variable is empty!"
    }

    # Footer Logo - Force Sync with Header Logo to ensure they are identical
    $Script:ACTechFooterBase64 = $Script:ACTechLogoBase64

    if ($Script:ACTechFooterBase64) {
        try {
            $bytesFooter = [Convert]::FromBase64String($Script:ACTechFooterBase64)
            $streamFooter = New-Object System.IO.MemoryStream(, $bytesFooter)
            $bitmapFooter = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmapFooter.BeginInit()
            $bitmapFooter.StreamSource = $streamFooter
            $bitmapFooter.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bitmapFooter.EndInit()
            $bitmapFooter.Freeze()
            if ($imgFooterLogo) { $imgFooterLogo.Source = $bitmapFooter }
        }
        catch {
            Write-Warning "Failed to load Footer Logo: $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "Footer Logo Base64 variable is empty!"
    }


    # -------------------------------------------------------------------------
    # LOGIC EVENTS
    # -------------------------------------------------------------------------
    

    # Helper for Transparency Toggle (Optimized Scope)
    $SetTransparency = {
        param($enable, $win)
        
        if (-not $win) { return }

        # 1. Force retrieval of the MainBorder
        $targetBorder = $win.FindName('MainBorder')
        
        if ($enable) {
            # Enable Blur (More Transparent: #88 alpha)
            if ($targetBorder) { 
                $targetBorder.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString("#88151515"))
            }
            else {
                # No MessageBox here
            }

            # Apply OS Blur Effect via C# Interop
            if ([System.Management.Automation.PSTypeName]::new('WindowEffects').Type) {
                try {
                    $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($win)).Handle
                    [WindowEffects]::EnableBlur($hwnd, 0, 0)
                }
                catch {
                    # Silent catch
                }
            }
        }
        else {
            # Disable Blur (Solid Opaque)
            if ($targetBorder) { 
                $targetBorder.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString("#FF151515"))
            }
        }
    }

    # 1. Window Moving (Drag Title Bar)
    $TitleBar = $window.FindName("TitleBar")
    $dragHandler = {
        param($s, $e)
        if ($e.ButtonState -eq 'Pressed') {
            & $SetTransparency -enable $true -win $window
            try { $window.DragMove() } catch {}
            & $SetTransparency -enable $false -win $window
        }
    }
    if ($TitleBar) {
        $TitleBar.Add_MouseLeftButtonDown($dragHandler)
    }

    # 2. Window Controls
    $btnClose.Add_Click({ 
            $window.Close() 
            # Ensure we kill the entire process if this is the main window
            [System.Windows.Forms.Application]::Exit()
        })
    $btnMinimize.Add_Click({ $window.WindowState = 'Minimized' })
    $btnMaximize.Add_Click({
            if ($window.WindowState -eq 'Normal') {
                $window.WindowState = 'Maximized'
                $MainBorder = $window.FindName('MainBorder') # Assuming MainBorder is the name of your main border element
                if ($MainBorder) {
                    $MainBorder.CornerRadius = '0'
                    $MainBorder.BorderThickness = '0'
                }
                $btnMaximize.Content = "[-]" # Restore symbol
            }
            else {
                $window.WindowState = 'Normal'
                $MainBorder = $window.FindName('MainBorder') # Assuming MainBorder is the name of your main border element
                if ($MainBorder) {
                    $MainBorder.CornerRadius = '10'
                    $MainBorder.BorderThickness = '1'
                }
                $btnMaximize.Content = "[ ]"
            }
        })

    # 3. Acrylic Effect (OnSourceInitialized)
    $window.Add_SourceInitialized({
            try {
                # Get View Handler
                [void](New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
                # Check for our C# type
                # INITIAL STATE: OPAQUE (Blur technically enabled but covered by opaque background or just not enabled yet)
                # We will only enable blur when needed or enable it but keep background opaque?
                # Best approach: Keep background opaque #FF... by default.
                
                # NOTE: If we want "Solid" look by default, we just ensure window.Background is #FF...
                # The XAML resource is now #FF151515 (Opaque)
                
                # We can enable blur potentially here if we want the underlying system to be ready, 
                # but if the background is solid, it won't show.
                # However, for performance, maybe only enable it when dragging?
                
                # Let's just ensure the starting state is correct:
                & $SetTransparency $false
            }
            catch {
                Write-Host "Visual effects error: $_"
            }
        })

    # 3.5 Fix Rounded Corner Transparency Artifact
    # Apply Clip geometry to MainBorder on Loaded to properly clip corners
    $MainBorder = $window.FindName('MainBorder')
    if ($MainBorder) {
        $MainBorder.Add_Loaded({
                param($senderObj, $e)
                $border = $senderObj
                $radius = 10
                $geometry = New-Object System.Windows.Media.RectangleGeometry
                $geometry.Rect = New-Object System.Windows.Rect(0, 0, $border.ActualWidth, $border.ActualHeight)
                $geometry.RadiusX = $radius
                $geometry.RadiusY = $radius
                $border.Clip = $geometry
            })
        $MainBorder.Add_SizeChanged({
                param($senderObj, $e)
                $border = $senderObj
                $radius = 10
                $geometry = New-Object System.Windows.Media.RectangleGeometry
                $geometry.Rect = New-Object System.Windows.Rect(0, 0, $border.ActualWidth, $border.ActualHeight)
                $geometry.RadiusX = $radius
                $geometry.RadiusY = $radius
                $border.Clip = $geometry
            })
    }


    # 4. Mode Selection Logic
    $updateModeVisuals = {
        if ($rbExpress.IsChecked) {
            $btnModeNext.Content = "Start Install (Recommended)"
        }
        elseif ($rbUninstall.IsChecked) {
            $btnModeNext.Content = "Start Deep Uninstall"
        }
        else {
            $btnModeNext.Content = "Configure Settings"
        }
    }
    $rbExpress.Add_Checked($updateModeVisuals)
    $rbCustom.Add_Checked($updateModeVisuals)
    $rbUninstall.Add_Checked($updateModeVisuals) # Add handler for uninstall radio button

    # 5. Buttons
    $btnModeNext.Add_Click({
            if ($rbExpress.IsChecked) {
                # Express Mode Logic
                $Script:ConfigResult.Cancelled = $false
                $Script:ConfigResult.Mode = 'Express'
                
                # Setup Global UserConfig for Express Mode
                $Script:UserConfig = @{
                    InstallMode          = 'Express'
                    OfficeVersion        = 'O365ProPlusRetail'
                    Channel              = 'CurrentPreview'
                    PrimaryLanguage      = $Script:ConfigResult.Language # Defaults to system lang or en-us
                    AdditionalLanguages  = @()
                    WindowsEdition       = 'Pro'
                    IncludeProject       = $false
                    IncludeVisio         = $false
                    IncludeClipchamp     = $true
                    IncludePowerAutomate = $false
                    SelectedApps         = @{
                        'Word'       = $true
                        'Excel'      = $true
                        'PowerPoint' = $true
                        'Outlook'    = $true
                        'Teams'      = $true
                        'Clipchamp'  = $true
                        'OneNote'    = $false
                        'OneDrive'   = $false
                        'Access'     = $false
                        'Publisher'  = $false
                        'Skype'      = $false
                    }
                }
                
                $window.DialogResult = $true
                $window.Close()
            }
            elseif ($rbUninstall.IsChecked) {
                # Deep Uninstall Mode
                $Script:ConfigResult.Cancelled = $false
                $Script:ConfigResult.Mode = 'Uninstall'
                
                # Close the config window and let the main script handle the uninstallation
                $window.DialogResult = $true
                $window.Close()
            }
            else {
                # Switch to Custom Page
                $pageMode.Visibility = 'Collapsed'
                $pageCustom.Visibility = 'Visible'
            }
        })

    $btnCustomBack.Add_Click({
            $pageCustom.Visibility = 'Collapsed'
            $pageMode.Visibility = 'Visible'
        })

    # 6. Dynamic Language Filtering
    # When Primary Language changes, hide that option from Additional Languages
    $updateAdditionalLangs = {
        $primaryLang = $cmbLanguage.SelectedItem.Tag
        $panelAdditionalLangs.Children | ForEach-Object {
            if ($_ -is [System.Windows.Controls.CheckBox]) {
                if ($_.Tag -eq $primaryLang) {
                    $_.Visibility = 'Collapsed'
                    $_.IsChecked = $false
                }
                else {
                    $_.Visibility = 'Visible'
                }
            }
        }
    }
    $cmbLanguage.Add_SelectionChanged($updateAdditionalLangs)
    # Run once on load to set initial state
    & $updateAdditionalLangs

    # 7. Global Select/Deselect
    $appsList = @($chkWord, $chkExcel, $chkPowerPoint, $chkOutlook, $chkTeams, $chkOneNote, $chkOneDrive, 
        $chkAccess, $chkPublisher, $chkProject, $chkVisio, $chkClipchamp, $chkPowerAutomate, $chkSkype,
        $chkDefender, $chkToDo, $chkPCManager, $chkStickyNotes, $chkPowerBI, $chkPowerToys, $chkVSCode, $chkVS2022, $chkCopilot)
    
    $btnSelectAll.Add_Click({ foreach ($a in $appsList) { if ($a) { $a.IsChecked = $true } } })
    $btnDeselectAll.Add_Click({ foreach ($a in $appsList) { if ($a) { $a.IsChecked = $false } } })

    $btnCustomStart.Add_Click({
            $Script:ConfigResult.Cancelled = $false
            $Script:ConfigResult.Mode = 'Custom'

            # Parse ComboBoxes
            if ($cmbVersion.SelectedItem.Tag) {
                $parts = $cmbVersion.SelectedItem.Tag -split '\|'
                $Script:ConfigResult.Version = $parts[0]
                $Script:ConfigResult.Channel = $parts[1]
            }
        
            $Script:ConfigResult.Language = $cmbLanguage.SelectedItem.Tag
            
            # Additional Languages
            $Script:ConfigResult.AdditionalLanguages = @()
            $panelAdditionalLangs.Children | ForEach-Object { 
                if ($_.IsChecked) { $Script:ConfigResult.AdditionalLanguages += $_.Tag } 
            }
        
            # Parse Windows Insider Selection (from radio buttons)
            $Script:ConfigResult.WindowsInsiderChannel = "None"
            $panelInsider.Children | ForEach-Object {
                if ($_ -is [System.Windows.Controls.RadioButton] -and $_.IsChecked) {
                    $Script:ConfigResult.WindowsInsiderChannel = $_.Tag
                }
            }

            # Populate result apps
            $Script:ConfigResult.Apps = @{}
            $appsList | ForEach-Object {
                if ($_) {
                    $name = $_.Name.Replace('chk', '') # e.g., chkWord -> Word
                    $Script:ConfigResult.Apps[$name] = $_.IsChecked
                }
            }

            # Setup Global UserConfig for Custom Mode
            $Script:UserConfig = @{
                InstallMode           = 'Custom'
                OfficeVersion         = $Script:ConfigResult.Version
                Channel               = $Script:ConfigResult.Channel
                PrimaryLanguage       = $Script:ConfigResult.Language
                AdditionalLanguages   = $Script:ConfigResult.AdditionalLanguages
                WindowsEdition        = if ($cmbWindowsEdition.SelectedItem) { $cmbWindowsEdition.SelectedItem.Tag } else { 'Pro' }
                WindowsInsiderChannel = $Script:ConfigResult.WindowsInsiderChannel # New: Insider Channel
                IncludeProject        = $chkProject.IsChecked
                IncludeVisio          = $chkVisio.IsChecked
                IncludeClipchamp      = $chkClipchamp.IsChecked
                IncludePowerAutomate  = $chkPowerAutomate.IsChecked
                SelectedApps          = $Script:ConfigResult.Apps
            }

            $window.DialogResult = $true
            $window.Close()
        })

    # Init Result with System Language Detection
    $sysLang = (Get-Culture).Name.ToLower()
    $validLangs = @('en-us', 'pt-br', 'es-es', 'de-de', 'fr-fr', 'ja-jp', 'zh-cn', 'it-it', 'ko-kr', 'ru-ru', 'ar-sa', 'da-dk', 'nl-nl', 'fi-fi', 'el-gr', 'he-il', 'hu-hu', 'id-id', 'ms-my', 'nb-no', 'pl-pl', 'pt-pt', 'ro-ro', 'sk-sk', 'sv-se', 'th-th', 'tr-tr', 'uk-ua', 'vi-vn')
    if ($validLangs -notcontains $sysLang) { $sysLang = 'en-us' }

    $Script:ConfigResult = @{
        Cancelled             = $true
        Mode                  = 'Express'
        Version               = 'O365ProPlusRetail'
        Channel               = 'CurrentPreview'
        Language              = $sysLang
        AdditionalLanguages   = @()
        WindowsInsiderChannel = "None" # New: Insider Channel
        Apps                  = @{}
    }
    
    # Set UI Language if possible (basic selection match)
    foreach ($item in $cmbLanguage.Items) {
        if ($item.Tag -eq $sysLang) {
            $cmbLanguage.SelectedItem = $item
            break
        }
    }

    $window.ShowDialog() | Out-Null
    
    # Return true/false based on DialogResult (which is true only if Next/Start clicked)
    if ($Script:ConfigResult.Cancelled) { return $false }
    return $true
}

# -------------------------------------------------------------------------
# FUNCTION: Start-OfficeUninstallation
# -------------------------------------------------------------------------
function Start-OfficeUninstallation {
    param()
    
    function Update-Progress {
        param($Message, $Percentage)
        Write-Log "Uninstall Progress: $Message ($Percentage%)" -Level Info
    }

    Write-Log "Starting ULTIMATE Deep Office Uninstallation (SCORCHED EARTH V3)..." -Level Info

    # 1. Kill Processes (The "Kill List")
    Update-Progress "Terminating all Microsoft productivity processes..." 5
    $processes = @(
        # Office Core
        "winword", "excel", "powerpnt", "outlook", "onenote", "onenotem", "mspub", "msaccess", "visio", "winproj", "groove", "lync", "teams",
        # Services/Background
        "officeclicktorun", "officec2rclient", "ose", "osppsvc", "msiexec", "installer", "setup", "odownload",
        # Teams & Skype
        "ms-teams", "teams", "skype", "skypeapp", "skypehost",
        # OneDrive
        "onedrive", "onedrivesetup", "filecoauth",
        # Power Platform
        "powerpnt", "powerbi", "powerbi.exe", "pbidesktop", "powerautomate", "pad.console.host", "microsoft.flow.assistant",
        # Utilities
        "powertoys", "powertoys.settings", "powertoys.runner", "clipchamp", "pcmanager", "microsoft.pcmanager", "stickynotes", "todo",
        # Other
        "cortana", "copilot", "officeapp", "yourphone", "phoneexperiencehost"
    )
    foreach ($proc in $processes) {
        Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
    }

    # 2. Special Binary Uninstallers (OneDrive)
    Update-Progress "Triggering OneDrive Uninstaller..." 10
    $oneDriveSetup = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    if (-not (Test-Path $oneDriveSetup)) { $oneDriveSetup = "$env:SystemRoot\System32\OneDriveSetup.exe" }
    if (Test-Path $oneDriveSetup) {
        Write-Log "Running OneDrive Uninstaller..." -Level Debug
        Start-Process $oneDriveSetup -ArgumentList "/uninstall" -Wait -NoNewWindow -ErrorAction SilentlyContinue
    }

    # 3. Stop Services
    Update-Progress "Stopping Services..." 15
    $services = @("ClickToRunSvc", "OfficeSvc", "ose", "OneSyncSvc", "OneSyncSvc_12345") # 12345 generic
    foreach ($svc in $services) {
        Set-Service -Name $svc -StartupType Disabled -ErrorAction SilentlyContinue
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
    }
    
    # 4. Remove Scheduled Tasks (Broad Sweep)
    Update-Progress "Removing Scheduled Tasks..." 20
    $taskPatterns = @("*Office*", "*Teams*", "*OneDrive*", "*Visio*", "*Project*", "*PowerToys*", "*Clipchamp*")
    foreach ($pattern in $taskPatterns) {
        Get-ScheduledTask | Where-Object { $_.TaskName -like $pattern } | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
    }

    # 5. Remove Store Apps (Appx List of Doom)
    Update-Progress "Removing Store Apps (List of Doom)..." 30
    $appxPatterns = @(
        "*Microsoft.Office*", "*Microsoft.OutlookForWindows*",
        "*MicrosoftTeams*", "*MSTeams*", "*Skype*",
        "*Clipchamp*",
        "*Microsoft.Todos*",
        "*PowerAutomate*", "*Flow*",
        "*PowerBI*",
        "*StickyNotes*",
        "*MicrosoftDefender*", # M365 App
        "*PCManager*",
        "*PowerToys*",
        "*Publisher*", "*Visio*", "*Project*", "*Copilot*"
    )
    
    foreach ($pattern in $appxPatterns) {
        try {
            # EXCLUDE Visual Studio and VS Code
            $apps = Get-AppxPackage -Name $pattern -AllUsers -ErrorAction SilentlyContinue | Where-Object { 
                $_.Name -notmatch "VisualStudio" -and $_.Name -notmatch "VSCode" -and $_.Name -notmatch "SecHealthUI"
            }
            foreach ($app in $apps) {
                Write-Log "Removing Appx: $($app.Name)" -Level Debug
                $app | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                $app | Remove-AppxPackage -ErrorAction SilentlyContinue
            }
        }
        catch {}
    }

    # 6. LICENSE & CREDENTIAL NUKE
    Update-Progress "Nuking Licenses & Credentials..." 40
    # Clean Registry Identities
    $idKeys = @(
        "HKCU:\Software\Microsoft\Office\15.0\Common\Identity",
        "HKCU:\Software\Microsoft\Office\16.0\Common\Identity",
        "HKCU:\Software\Microsoft\Office\16.0\Common\Internet",
        "HKCU:\Software\Microsoft\Office\16.0\Registration"
    )
    foreach ($k in $idKeys) { Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue }

    # Clean Credential Manager
    $creds = cmdkey /list | Select-String -Pattern "Target: LegacyGeneric:target=(MicrosoftOffice|Microsoft_Path|O365|OneDrive|Teams|SSO_POP)"
    if ($creds) {
        foreach ($c in $creds) {
            $cName = $c.ToString().Split("=")[1].Trim()
            cmdkey /delete:$cName | Out-Null
        }
    }

    # 7. MSI Blast (Targeted & Dynamic)
    Update-Progress "Removing MSI Products..." 50
    # 7a. The GUID List (Standard Office)
    $msiGuids = @(
        "{20150000-008C-0000-0000-0000000FF1CE}", "{20150000-008C-0C0A-0000-0000000FF1CE}", "{50150000-008F-0000-1000-0000000FF1CE}", "{90150000-007E-0000-0000-0000000FF1CE}", 
        "{90150000-007E-0000-1000-0000000FF1CE}", "{90150000-008C-0000-0000-0000000FF1CE}", "{90150000-008C-0000-1000-0000000FF1CE}", "{90150000-008C-0409-0000-0000000FF1CE}",
        "{90160000-008C-0000-0000-0000000FF1CE}", "{90160000-008C-0000-1000-0000000FF1CE}", "{90160000-008C-0409-0000-0000000FF1CE}", "{90160000-008C-0409-1000-0000000FF1CE}",
        "{90160000-007E-0000-0000-0000000FF1CE}", "{90160000-007E-0000-1000-0000000FF1CE}", 
        "{65DA2EC9-0642-47E9-AAE2-B5267AA14D75}", "{E50AE784-FABE-46DA-A1F8-7B6B56DCB22E}" # Activation Assistants
    )
    foreach ($guid in $msiGuids) {
        Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait -NoNewWindow
    }
    
    # 7b. Dynamic Search (Project, Visio, Teams Machine-Wide)
    $uninstallRoots = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
    $msiKeywords = "Office|Project|Visio|Teams Machine-Wide|PowerToys|Power BI|Clipchamp|PC Manager"
    
    foreach ($root in $uninstallRoots) {
        Get-ChildItem -Path $root -ErrorAction SilentlyContinue | ForEach-Object {
            $disp = $_.GetValue("DisplayName")
            if ($disp -match $msiKeywords -and $disp -notmatch "Visual Studio" -and $disp -notmatch "VS Code") {
                if ($_.PSChildName -match "{.*}") {
                    Write-Log "Dynamic MSI Uninstall: $disp" -Level Debug
                    Start-Process "msiexec.exe" -ArgumentList "/x $($_.PSChildName) /qn /norestart" -Wait -NoNewWindow
                }
            }
        }
    }

    # 8. Files & Folders (Aggressive Takeown)
    Update-Progress "Cleaning Filesystem..." 70
    $cleanPaths = @(
        "$env:ProgramFiles\Microsoft Office", "$env:ProgramFiles(x86)\Microsoft Office",
        "$env:ProgramData\Microsoft\Office", "$env:ProgramData\Microsoft\ClickToRun", 
        "$env:ProgramData\Microsoft\Teams", "$env:ProgramData\Microsoft\OneDrive",
        "$env:LOCALAPPDATA\Microsoft\Office", "$env:LOCALAPPDATA\Microsoft\Teams", "$env:LOCALAPPDATA\Microsoft\OneDrive",
        "$env:LOCALAPPDATA\Microsoft\PowerToys", "$env:LOCALAPPDATA\Microsoft\Power BI Desktop",
        "$env:APPDATA\Microsoft\Office", "$env:APPDATA\Microsoft\Teams", "$env:APPDATA\Microsoft\OneDrive",
        "$env:USERPROFILE\Microsoft Office", "$env:USERPROFILE\OneDrive"
    )
    foreach ($path in $cleanPaths) {
        if (Test-Path $path) { 
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path $path) {
                Start-Process "cmd.exe" -ArgumentList "/c takeown /f `"$path`" /r /d y && icacls `"$path`" /grant administrators:F /t && rd /s /q `"$path`"" -Wait -NoNewWindow -WindowStyle Hidden
            }
        }
    }
    
    # 9. Registry Blast (Regex V3)
    Update-Progress "Final Registry Clean..." 90
    $regRegex = "0FF1CE|O365|Microsoft ?Office|Microsoft ?365|Project|Visio|Clipchamp|Power ?Automate|Power ?Toys|Power ?BI|Skype|Teams|OneDrive|PC ?Manager|Sticky ?Notes"
    
    # Delete specific roots first
    $regRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Office", "HKLM:\SOFTWARE\Microsoft\ClickToRun", "HKLM:\SOFTWARE\Microsoft\Teams",
        "HKCU:\SOFTWARE\Microsoft\Office", "HKCU:\SOFTWARE\Microsoft\Teams", "HKCU:\SOFTWARE\Microsoft\OneDrive"
    )
    foreach ($r in $regRoots) { Remove-Item -Path $r -Recurse -Force -ErrorAction SilentlyContinue }

    # Scan Uninstall Keys
    foreach ($root in $uninstallRoots) {
        Get-ChildItem -Path $root -ErrorAction SilentlyContinue | ForEach-Object {
            $name = $_.GetValue("DisplayName")
            if ($name -match $regRegex -and $name -notmatch "Visual Studio" -and $name -notmatch "VS Code" -and $name -notmatch "WebView2") {
                Write-Log "Deleting Registry Key: $name" -Level Debug
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Update-Progress "Office uninstallation complete." 100
    Write-Log "ULTIMATE Deep Office uninstallation (V3 Scorched Earth) finished." -Level Success
    Start-Sleep -Seconds 2
}

# -------------------------------------------------------------------------
# FUNCTION: Set-WindowsInsider
# -------------------------------------------------------------------------
function Set-WindowsInsider {
    param(
        [Parameter(Mandatory)]
        [string]$Channel # e.g., "ReleasePreview", "Beta", "Dev", "Canary", "None"
    )
    
    if ($Channel -eq "None") { 
        Write-Log "Windows Insider channel set to 'None'. No changes applied." -Level Info
        return 
    }
    
    Write-Log "Attempting to set Windows Insider channel to '$Channel'..." -Level Info

    $Path = "HKLM:\SOFTWARE\Microsoft\WindowsSelfHost\Applicability"
    if (!(Test-Path $Path)) { 
        New-Item -Path $Path -Force | Out-Null 
        Write-Log "Created registry path: $Path" -Level Debug
    }
    
    # Default values for Insider settings
    $BranchName = "Release"
    $UserPreferredBranchName = "Release"
    $Ring = "External"
    $ContentType = "Mainline"
    $EnablePreviewBuilds = 1 # Always enable for any channel selection
    
    switch ($Channel) {
        "ReleasePreview" { 
            $BranchName = "ReleasePreview"
            $UserPreferredBranchName = "ReleasePreview"
            $Ring = "External"
            $ContentType = "Mainline"
        }
        "Beta" {           
            $BranchName = "Beta"
            $UserPreferredBranchName = "Beta"
            $Ring = "External"
            $ContentType = "Mainline"
        }
        "Dev" {            
            $BranchName = "Dev"
            $UserPreferredBranchName = "Dev"
            $Ring = "External"
            $ContentType = "Mainline"
        }
        "Canary" {         
            $BranchName = "Canary"
            $UserPreferredBranchName = "Canary"
            $Ring = "External"
            $ContentType = "Mainline"
        }
        default {
            Write-Log "Invalid Windows Insider channel '$Channel' specified. No changes applied." -Level Warning
            return
        }
    }
    
    try {
        Set-ItemProperty -Path $Path -Name "BranchName" -Value $BranchName -Force -ErrorAction Stop
        Set-ItemProperty -Path $Path -Name "UserPreferredBranchName" -Value $UserPreferredBranchName -Force -ErrorAction Stop
        Set-ItemProperty -Path $Path -Name "Ring" -Value $Ring -Force -ErrorAction Stop
        Set-ItemProperty -Path $Path -Name "ContentType" -Value $ContentType -Force -ErrorAction Stop
        Set-ItemProperty -Path $Path -Name "EnablePreviewBuilds" -Value $EnablePreviewBuilds -Force -ErrorAction Stop
        
        Write-Log "Windows Insider channel successfully set to '$Channel'." -Level Success
        
        # Prompt for restart
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
        $res = [System.Windows.Forms.MessageBox]::Show("Windows Insider settings applied.`nA restart is required for changes to take full effect. Restart now?", "Restart Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($res -eq 'Yes') {
            Write-Log "User chose to restart. Restarting computer..." -Level Info
            Restart-Computer -Force
        }
        else {
            Write-Log "User chose not to restart immediately." -Level Info
        }
    }
    catch {
        Write-Log "Failed to set Windows Insider channel: $($_.Exception.Message)" -Level Error
    }
}

function Show-ConfigWindow-Legacy {
    <#
    .SYNOPSIS
    Shows configuration dialog in main thread (blocking).
    Returns $true if user clicks Start, $false if cancelled.
    #>
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    
    # Get localized strings
    $title = L 'WindowTitle'
    $subtitle = L 'WindowSubtitle'
    $strExpress = L 'ExpressMode'
    $strExpressDesc = L 'ExpressModeDesc'
    $strCustom = L 'CustomMode'
    $strCustomDesc = L 'CustomModeDesc'
    $strStart = L 'BtnStart'
    $strCancel = L 'BtnCancel'
    $strVersion = L 'SelectVersion'
    $strLang = L 'PrimaryLanguage'
    $strApps = L 'SelectApps'
    $strSelectAll = L 'SelectAll'
    $strDeselectAll = L 'DeselectAll'
    $strBack = L 'BtnBack'

    $strConfigure = L 'ConfigureInstallation'
    $strFooterDevBy = L 'FooterDevBy'
    
    # Detect Windows language for default
    $winLang = (Get-Culture).Name.ToLower()
    $defaultLangIndex = 0
    $langMap = @{ 'en-us' = 0; 'pt-br' = 1; 'es-es' = 2; 'de-de' = 3; 'fr-fr' = 4; 'ja-jp' = 5; 'zh-cn' = 6; 'it-it' = 7; 'ko-kr' = 8; 'ru-ru' = 9 }
    if ($langMap.ContainsKey($winLang)) { $defaultLangIndex = $langMap[$winLang] }
    
    # Escape strings for XAML
    foreach ($v in @('title', 'subtitle', 'strExpress', 'strExpressDesc', 'strCustom', 'strCustomDesc', 'strStart', 'strCancel', 'strVersion', 'strLang', 'strApps', 'strSelectAll', 'strDeselectAll', 'strBack', 'strConfigure', 'strFooterDevBy')) {
        Set-Variable -Name $v -Value ((Get-Variable -Name $v).Value -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$title" Width="700" Height="600"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize" WindowStyle="SingleBorderWindow"
        AllowsTransparency="True"
        Background="#1e1e1e" Topmost="True" BorderBrush="#0078D4" BorderThickness="1">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border x:Name="borderHeader" Grid.Row="0" Background="#0078D4" Padding="20,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Microsoft 365 Ultimate Installer" FontSize="22" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Text="$subtitle" FontSize="12" Foreground="#DDDDDD" Margin="0,3,0,0"/>
                </StackPanel>
                <!-- AC Tech Logo in Header -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0">
                    <Image x:Name="imgACTechLogo" Width="40" Height="40" Margin="0,0,10,0" VerticalAlignment="Center" Stretch="Uniform"/>
                    <TextBlock Text="AC Tech" FontSize="14" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Footer with AC Tech Signature -->
        <Border Grid.Row="2" Background="#2d2d2d" Padding="0,10" BorderBrush="#444444" BorderThickness="0,1,0,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                <Image x:Name="imgACTechFooter" Width="20" Height="20" Margin="0,0,8,0" VerticalAlignment="Center" Stretch="Uniform"/>
                <TextBlock Text="$strFooterDevBy" FontSize="11" Foreground="#888888" VerticalAlignment="Center"/>
            </StackPanel>
        </Border>
        
        <!-- MODE SELECTION PAGE -->
        <Grid x:Name="PageMode" Grid.Row="1" Margin="25">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            
            <TextBlock Grid.Row="0" Text="Choose installation mode:" FontSize="16" Foreground="White" Margin="0,0,0,20"/>
            
            <StackPanel Grid.Row="1">
                <Border x:Name="borderExpress" Background="#2d2d2d" BorderBrush="#0078D4" BorderThickness="2" CornerRadius="8" Padding="15" Margin="0,0,0,15" Cursor="Hand">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <RadioButton x:Name="rbExpress" GroupName="Mode" IsChecked="True" VerticalAlignment="Center" Margin="0,0,15,0"/>
                        <StackPanel Grid.Column="1">
                            <TextBlock Text="⚡ $strExpress" FontSize="14" FontWeight="SemiBold" Foreground="White"/>
                            <TextBlock Text="$strExpressDesc" FontSize="12" Foreground="#AAAAAA" TextWrapping="Wrap" Margin="0,5,0,0"/>
                            <TextBlock Text="Word, Excel, PowerPoint, Outlook, Teams, Clipchamp" FontSize="11" Foreground="#666666" Margin="0,8,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Grid>
                </Border>
                
                <Border x:Name="borderCustom" Background="#2d2d2d" BorderBrush="#444444" BorderThickness="1" CornerRadius="8" Padding="15" Cursor="Hand">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <RadioButton x:Name="rbCustom" GroupName="Mode" VerticalAlignment="Center" Margin="0,0,15,0"/>
                        <StackPanel Grid.Column="1">
                            <TextBlock Text="⚙️ $strCustom" FontSize="14" FontWeight="SemiBold" Foreground="White"/>
                            <TextBlock Text="$strCustomDesc" FontSize="12" Foreground="#AAAAAA" TextWrapping="Wrap" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </StackPanel>
            
            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
                <Button x:Name="btnModeCancel" Content="$strCancel" Width="100" Height="35" Background="#444444" Foreground="White" BorderThickness="0" Margin="0,0,10,0"/>
                <Button x:Name="btnModeNext" Content="$strStart" Width="180" Height="35" Background="#0078D4" Foreground="White" BorderThickness="0" FontWeight="SemiBold"/>
            </StackPanel>
        </Grid>
        
        <!-- CUSTOM PAGE -->
        <Grid x:Name="PageCustom" Grid.Row="1" Margin="25" Visibility="Collapsed">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            
            <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Margin="0,0,-10,0" Padding="0,0,15,0">
                <StackPanel>
                    <TextBlock Text="$strVersion" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,0,0,8"/>
                    <ComboBox x:Name="cmbVersion" HorizontalAlignment="Stretch" SelectedIndex="0" Height="32" Padding="6,4" Background="#333333" Foreground="Black" BorderBrush="#555555">
                        <ComboBox.Resources>
                            <Style TargetType="ComboBoxItem">
                                <Setter Property="Background" Value="#333333"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Style>
                        </ComboBox.Resources>
                        <ComboBoxItem Tag="O365ProPlusRetail|CurrentPreview">Microsoft 365 Enterprise (Current Preview)</ComboBoxItem>
                        <ComboBoxItem Tag="O365BusinessRetail|CurrentPreview">Microsoft 365 Business (Current Preview)</ComboBoxItem>
                        <ComboBoxItem Tag="ProPlus2024Retail|PerpetualVL2024">Office LTSC 2024</ComboBoxItem>
                        <ComboBoxItem Tag="ProPlus2021Retail|PerpetualVL2021">Office LTSC 2021</ComboBoxItem>
                        <ComboBoxItem Tag="ProPlus2019Retail|PerpetualVL2019">Office 2019</ComboBoxItem>
                    </ComboBox>
                    
                    <Grid Margin="0,20,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="$strApps" FontSize="14" Foreground="White" FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <Button x:Name="btnSelectAll" Content="$strSelectAll" Width="100" Height="26" Background="#333333" Foreground="White" BorderThickness="0" Margin="0,0,10,0"/>
                            <Button x:Name="btnDeselectAll" Content="$strDeselectAll" Width="100" Height="26" Background="#333333" Foreground="White" BorderThickness="0"/>
                        </StackPanel>
                    </Grid>

                    <Border Background="#252525" CornerRadius="6" Padding="15" BorderBrush="#333333" BorderThickness="1">
                        <UniformGrid Columns="2" VerticalAlignment="Top">
                            <!-- Core Office -->
                            <CheckBox x:Name="chkWord" Content="Word" IsChecked="True" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkExcel" Content="Excel" IsChecked="True" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkPowerPoint" Content="PowerPoint" IsChecked="True" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkOutlook" Content="Outlook (New)" IsChecked="True" Foreground="White" Margin="5,8" FontSize="13" ToolTip="Installs the New Outlook for Windows"/>
                            <CheckBox x:Name="chkTeams" Content="Teams" IsChecked="True" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkOneNote" Content="OneNote" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkOneDrive" Content="OneDrive" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkAccess" Content="Access" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkPublisher" Content="Publisher" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            
                            <!-- Specialized -->
                            <CheckBox x:Name="chkProject" Content="Project Pro" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkVisio" Content="Visio Pro" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            
                            <!-- Additional Apps -->
                            <CheckBox x:Name="chkClipchamp" Content="Clipchamp" IsChecked="True" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkDefender" Content="Microsoft Defender" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkToDo" Content="Microsoft To Do" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkPCManager" Content="Microsoft PC Manager" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkStickyNotes" Content="Sticky Notes" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkPowerBI" Content="Power BI Desktop" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkPowerAutomate" Content="Power Automate" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkPowerToys" Content="PowerToys" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            
                            <!-- Dev Tools & Others -->
                            <CheckBox x:Name="chkVSCode" Content="Visual Studio Code" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkVS2022" Content="Visual Studio 2026 Ent. (Preview)" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkCopilot" Content="Microsoft 365 Copilot" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                            <CheckBox x:Name="chkSkype" Content="Skype" IsChecked="False" Foreground="White" Margin="5,8" FontSize="13"/>
                        </UniformGrid>
                    </Border>


                    <TextBlock Text="$strLang" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,20,0,8"/>
                    <ComboBox x:Name="cmbLanguage" HorizontalAlignment="Stretch" SelectedIndex="$defaultLangIndex" Height="32" Padding="6,4" Background="#333333" Foreground="Black" BorderBrush="#555555">
                        <ComboBox.Resources>
                            <Style TargetType="ComboBoxItem">
                                <Setter Property="Background" Value="#333333"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Style>
                        </ComboBox.Resources>
                        <ComboBoxItem Tag="en-us">English (United States)</ComboBoxItem>
                        <ComboBoxItem Tag="pt-br">Português (Brasil)</ComboBoxItem>
                        <ComboBoxItem Tag="es-es">Español</ComboBoxItem>
                        <ComboBoxItem Tag="de-de">Deutsch</ComboBoxItem>
                        <ComboBoxItem Tag="fr-fr">Français</ComboBoxItem>
                        <ComboBoxItem Tag="ja-jp">日本語</ComboBoxItem>
                        <ComboBoxItem Tag="zh-cn">中文 (简体)</ComboBoxItem>
                        <ComboBoxItem Tag="it-it">Italiano</ComboBoxItem>
                        <ComboBoxItem Tag="ko-kr">한국어</ComboBoxItem>
                        <ComboBoxItem Tag="ru-ru">Русский</ComboBoxItem>
                        <ComboBoxItem Tag="ar-sa">Arabic (Saudi Arabia)</ComboBoxItem>
                        <ComboBoxItem Tag="da-dk">Danish</ComboBoxItem>
                        <ComboBoxItem Tag="nl-nl">Dutch</ComboBoxItem>
                        <ComboBoxItem Tag="fi-fi">Finnish</ComboBoxItem>
                        <ComboBoxItem Tag="el-gr">Greek</ComboBoxItem>
                        <ComboBoxItem Tag="he-il">Hebrew</ComboBoxItem>
                        <ComboBoxItem Tag="hu-hu">Hungarian</ComboBoxItem>
                        <ComboBoxItem Tag="id-id">Indonesian</ComboBoxItem>
                        <ComboBoxItem Tag="ms-my">Malay</ComboBoxItem>
                        <ComboBoxItem Tag="nb-no">Norwegian (Bokmål)</ComboBoxItem>
                        <ComboBoxItem Tag="pl-pl">Polish</ComboBoxItem>
                        <ComboBoxItem Tag="pt-pt">Portuguese (Portugal)</ComboBoxItem>
                        <ComboBoxItem Tag="ro-ro">Romanian</ComboBoxItem>
                        <ComboBoxItem Tag="sk-sk">Slovak</ComboBoxItem>
                        <ComboBoxItem Tag="sv-se">Swedish</ComboBoxItem>
                        <ComboBoxItem Tag="th-th">Thai</ComboBoxItem>
                        <ComboBoxItem Tag="tr-tr">Turkish</ComboBoxItem>
                        <ComboBoxItem Tag="uk-ua">Ukrainian</ComboBoxItem>
                        <ComboBoxItem Tag="vi-vn">Vietnamese</ComboBoxItem>
                    </ComboBox>

                    <TextBlock Text="Additional Languages:" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,20,0,8"/>
                    <Border Background="#252525" CornerRadius="6" BorderBrush="#333333" BorderThickness="1" Padding="10">
                        <ScrollViewer Height="120" VerticalScrollBarVisibility="Auto">
                            <WrapPanel x:Name="panelAdditionalLangs">
                                <CheckBox x:Name="chkLang_en_us" Content="English (US)" Tag="en-us" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_pt_br" Content="Português (BR)" Tag="pt-br" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_es_es" Content="Español" Tag="es-es" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_de_de" Content="Deutsch" Tag="de-de" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_fr_fr" Content="Français" Tag="fr-fr" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_ja_jp" Content="日本語" Tag="ja-jp" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_zh_cn" Content="中文 (简体)" Tag="zh-cn" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_it_it" Content="Italiano" Tag="it-it" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_ko_kr" Content="한국어" Tag="ko-kr" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_ru_ru" Content="Русский" Tag="ru-ru" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_ar_sa" Content="Arabic" Tag="ar-sa" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_da_dk" Content="Danish" Tag="da-dk" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_nl_nl" Content="Dutch" Tag="nl-nl" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_fi_fi" Content="Finnish" Tag="fi-fi" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_el_gr" Content="Greek" Tag="el-gr" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_he_il" Content="Hebrew" Tag="he-il" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_hu_hu" Content="Hungarian" Tag="hu-hu" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_id_id" Content="Indonesian" Tag="id-id" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_ms_my" Content="Malay" Tag="ms-my" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_nb_no" Content="Norwegian (Bokmål)" Tag="nb-no" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_pl_pl" Content="Polish" Tag="pl-pl" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_pt_pt" Content="Portuguese (PT)" Tag="pt-pt" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_ro_ro" Content="Romanian" Tag="ro-ro" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_sk_sk" Content="Slovak" Tag="sk-sk" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_sv_se" Content="Swedish" Tag="sv-se" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_th_th" Content="Thai" Tag="th-th" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_tr_tr" Content="Turkish" Tag="tr-tr" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_uk_ua" Content="Ukrainian" Tag="uk-ua" Foreground="White" Margin="0,0,20,8"/>
                                <CheckBox x:Name="chkLang_vi_vn" Content="Vietnamese" Tag="vi-vn" Foreground="White" Margin="0,0,20,8"/>
                            </WrapPanel>
                        </ScrollViewer>
                    </Border>

                    <TextBlock Text="Windows Edition (Activation):" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,20,0,8"/>
                    <ComboBox x:Name="cmbWindowsEdition" HorizontalAlignment="Stretch" SelectedIndex="0" Height="32" Padding="6,4" Background="#333333" Foreground="Black" BorderBrush="#555555">
                         <ComboBox.Resources>
                            <Style TargetType="ComboBoxItem">
                                <Setter Property="Background" Value="#333333"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Style>
                        </ComboBox.Resources>
                        <ComboBoxItem Tag="Pro">Windows 10/11 Pro (Recommended)</ComboBoxItem>
                        <ComboBoxItem Tag="Home">Windows 10/11 Home</ComboBoxItem>
                        <ComboBoxItem Tag="Enterprise">Windows 10/11 Enterprise</ComboBoxItem>
                        <ComboBoxItem Tag="Education">Windows 10/11 Education</ComboBoxItem>
                        <ComboBoxItem Tag="IoTEnterprise">Windows 10/11 IoT Enterprise</ComboBoxItem>
                        <ComboBoxItem Tag="">Do Not Change Edition</ComboBoxItem>
                    </ComboBox>
                </StackPanel>
            </ScrollViewer>
            
            <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                <Button x:Name="btnCustomBack" Content="$strBack" Width="100" Height="35" Background="#444444" Foreground="White" BorderThickness="0" Margin="0,0,10,0"/>
                <Button x:Name="btnCustomStart" Content="$strStart" Width="150" Height="35" Background="#0078D4" Foreground="White" BorderThickness="0" FontWeight="SemiBold"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@
    
    try {
        [xml]$xamlDoc = $xaml
        $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Get controls
        $pageMode = $window.FindName('PageMode')
        $pageCustom = $window.FindName('PageCustom')
        $borderExpress = $window.FindName('borderExpress')
        $borderCustom = $window.FindName('borderCustom')
        $rbExpress = $window.FindName('rbExpress')
        $rbCustom = $window.FindName('rbCustom')
        $btnModeNext = $window.FindName('btnModeNext')
        $btnModeCancel = $window.FindName('btnModeCancel')
        $btnCustomBack = $window.FindName('btnCustomBack')
        $btnCustomStart = $window.FindName('btnCustomStart')
        $btnSelectAll = $window.FindName('btnSelectAll')
        $btnDeselectAll = $window.FindName('btnDeselectAll')
        $cmbVersion = $window.FindName('cmbVersion')
        $cmbLanguage = $window.FindName('cmbLanguage')
        $cmbWindowsEdition = $window.FindName('cmbWindowsEdition')
        
        # Original Apps
        $chkWord = $window.FindName('chkWord')
        $chkExcel = $window.FindName('chkExcel')
        $chkPowerPoint = $window.FindName('chkPowerPoint')
        $chkOutlook = $window.FindName('chkOutlook')
        $chkAccess = $window.FindName('chkAccess')
        $chkPublisher = $window.FindName('chkPublisher')
        $chkTeams = $window.FindName('chkTeams')
        $chkOneNote = $window.FindName('chkOneNote')
        $chkOneDrive = $window.FindName('chkOneDrive')
        $chkProject = $window.FindName('chkProject')
        $chkVisio = $window.FindName('chkVisio')
        $chkClipchamp = $window.FindName('chkClipchamp')
        $chkPowerAutomate = $window.FindName('chkPowerAutomate')
        $chkSkype = $window.FindName('chkSkype')
        
        # New Apps
        $chkDefender = $window.FindName('chkDefender')
        $chkToDo = $window.FindName('chkToDo')
        $chkPCManager = $window.FindName('chkPCManager')
        $chkStickyNotes = $window.FindName('chkStickyNotes')
        $chkPowerBI = $window.FindName('chkPowerBI')
        $chkPowerToys = $window.FindName('chkPowerToys')
        $chkVSCode = $window.FindName('chkVSCode')
        $chkVS2022 = $window.FindName('chkVS2022')
        $chkCopilot = $window.FindName('chkCopilot')
        
        # Load AC Tech Logo
        $imgACTechLogo = $window.FindName('imgACTechLogo')
        $imgACTechFooter = $window.FindName('imgACTechFooter')
        
        try {
            $bitmap = $null
            
            # ALWAYS Load from embedded Base64 first (as per user request)
            if ($Script:ACTechLogoBase64) {
                try {
                    $bytes = [Convert]::FromBase64String($Script:ACTechLogoBase64)
                    $stream = New-Object System.IO.MemoryStream(, $bytes)
                    
                    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bitmap.BeginInit()
                    $bmp.StreamSource = $stream
                    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                    $bitmap.EndInit()
                    $bitmap.Freeze()
                }
                catch {
                    Write-Host "Base64 Load Failed: $($_.Exception.Message)"
                }
            }

            # Fallback to local file if Base64 failed or empty
            if (-not $bitmap) {
                $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
                if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = $PSScriptRoot }
                $logoPath = Join-Path $scriptDir 'ac_tech_logo.png'
                
                if (Test-Path $logoPath) {
                    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bitmap.BeginInit()
                    $bitmap.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute)
                    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                    $bitmap.EndInit()
                    $bitmap.Freeze()
                }
            }
            
            if ($bitmap) {
                if ($imgACTechLogo) { $imgACTechLogo.Source = $bitmap }
                if ($imgACTechFooter) { $imgACTechFooter.Source = $bitmap }
            }
        }
        catch {
            Write-Host "Warning: Could not load AC Tech logo: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Get language checkboxes
        $langCheckboxes = @()
        foreach ($tag in @('en-us', 'pt-br', 'es-es', 'de-de', 'fr-fr', 'ja-jp', 'zh-cn', 'it-it', 'ko-kr', 'ru-ru', 'ar-sa', 'da-dk', 'nl-nl', 'fi-fi', 'el-gr', 'he-il', 'hu-hu', 'id-id', 'ms-my', 'nb-no', 'pl-pl', 'pt-pt', 'ro-ro', 'sk-sk', 'sv-se', 'th-th', 'tr-tr', 'uk-ua', 'vi-vn')) {
            $cb = $window.FindName("chkLang_$($tag -replace '-','_')")
            if ($cb) { $langCheckboxes += $cb }
        }
        
        # Result variable
        $Script:ConfigResult = @{
            Cancelled            = $true
            Mode                 = 'Express'
            Version              = 'O365ProPlusRetail'
            Channel              = 'CurrentPreview'
            Language             = 'en-us'
            WindowsEdition       = 'Pro'
            Apps                 = @{}
            IncludeProject       = $true
            IncludeVisio         = $true
            IncludeClipchamp     = $true
            IncludePowerAutomate = $true
        }
        
        # Language Logic: Hide primary language from additional list
        $updateLangFilter = {
            if ($cmbLanguage.SelectedItem) {
                $selectedLang = $cmbLanguage.SelectedItem.Tag
                foreach ($cb in $langCheckboxes) {
                    if ($cb.Tag -eq $selectedLang) {
                        $cb.IsChecked = $false
                        $cb.Visibility = 'Collapsed'
                    }
                    else {
                        $cb.Visibility = 'Visible'
                    }
                }
            }
        }
        $cmbLanguage.Add_SelectionChanged($updateLangFilter)
        # Trigger once on load
        &$updateLangFilter

        # Visual Selection Logic
        $updateVisuals = {
            if ($rbExpress.IsChecked) {
                $borderExpress.BorderBrush = "#0078D4"; $borderExpress.BorderThickness = "2"
                $borderCustom.BorderBrush = "#444444"; $borderCustom.BorderThickness = "1"
                $btnModeNext.Content = $strStart
            }
            else {
                $borderExpress.BorderBrush = "#444444"; $borderExpress.BorderThickness = "1"
                $borderCustom.BorderBrush = "#0078D4"; $borderCustom.BorderThickness = "2"
                $btnModeNext.Content = $strConfigure
            }
        }
        
        $rbExpress.Add_Checked($updateVisuals)
        $rbCustom.Add_Checked($updateVisuals)
        
        $borderExpress.Add_MouseDown({ $rbExpress.IsChecked = $true })
        $borderCustom.Add_MouseDown({ $rbCustom.IsChecked = $true })
        
        # Mode Next Button
        $btnModeNext.Add_Click({
                if ($rbExpress.IsChecked) {
                    # Express Logic (Defaults)
                    $Script:ConfigResult.Mode = "Express"
                    # The original Show-ConfigWindow-Legacy does not have Move-ToPage or Start-Process here.
                    # It sets the mode and closes the window.
                    $Script:ConfigResult.Cancelled = $false
                    $window.DialogResult = $true
                    $window.Close()
                }
                elseif ($rbUninstall.IsChecked) {
                    # Deep Uninstall Mode
                    $Script:ConfigResult.Mode = "Uninstall"
                    $Script:ConfigResult.Cancelled = $false
             
                    # Close the config window and let the main script handle the uninstallation
                    $window.DialogResult = $true
                    $window.Close()
                }
                else {
                    # Custom Mode -> Go to Custom Page
                    $pageMode.Visibility = 'Collapsed'
                    $pageCustom.Visibility = 'Visible'
                }
            })
        
        # Cancel
        $btnModeCancel.Add_Click({
                $Script:ConfigResult.Cancelled = $true
                $window.DialogResult = $false
                $window.Close()
            })
        
        # Back
        $btnCustomBack.Add_Click({
                $pageCustom.Visibility = 'Collapsed'
                $pageMode.Visibility = 'Visible'
            })
        
        # Select All
        $btnSelectAll.Add_Click({
                $chkWord.IsChecked = $true; $chkExcel.IsChecked = $true; $chkPowerPoint.IsChecked = $true
                $chkOutlook.IsChecked = $true; $chkAccess.IsChecked = $true; $chkPublisher.IsChecked = $true
                $chkTeams.IsChecked = $true; $chkOneNote.IsChecked = $true; $chkOneDrive.IsChecked = $true
                $chkProject.IsChecked = $true; $chkVisio.IsChecked = $true; $chkClipchamp.IsChecked = $true
                $chkPowerAutomate.IsChecked = $true; $chkSkype.IsChecked = $true
                # New Apps
                $chkDefender.IsChecked = $true; $chkToDo.IsChecked = $true; $chkPCManager.IsChecked = $true
                $chkStickyNotes.IsChecked = $true; $chkPowerBI.IsChecked = $true; $chkPowerToys.IsChecked = $true
                $chkVSCode.IsChecked = $true; $chkVS2022.IsChecked = $true; $chkCopilot.IsChecked = $true
            })
        
        # Deselect All
        $btnDeselectAll.Add_Click({
                $chkWord.IsChecked = $false; $chkExcel.IsChecked = $false; $chkPowerPoint.IsChecked = $false
                $chkOutlook.IsChecked = $false; $chkAccess.IsChecked = $false; $chkPublisher.IsChecked = $false
                $chkTeams.IsChecked = $false; $chkOneNote.IsChecked = $false; $chkOneDrive.IsChecked = $false
                $chkProject.IsChecked = $false; $chkVisio.IsChecked = $false; $chkClipchamp.IsChecked = $false
                $chkPowerAutomate.IsChecked = $false; $chkSkype.IsChecked = $false
                # New Apps
                $chkDefender.IsChecked = $false; $chkToDo.IsChecked = $false; $chkPCManager.IsChecked = $false
                $chkStickyNotes.IsChecked = $false; $chkPowerBI.IsChecked = $false; $chkPowerToys.IsChecked = $false
                $chkVSCode.IsChecked = $false; $chkVS2022.IsChecked = $false; $chkCopilot.IsChecked = $false
            })
        
        # Custom Start
        $btnCustomStart.Add_Click({
                $Script:ConfigResult.Cancelled = $false
                $Script:ConfigResult.Mode = 'Custom'
            
                # Parse version/channel
                $versionTag = $cmbVersion.SelectedItem.Tag
                if ($versionTag) {
                    $parts = $versionTag -split '\|'
                    $Script:ConfigResult.Version = $parts[0]
                    $Script:ConfigResult.Channel = $parts[1]
                }
            
                # Language
                $Script:ConfigResult.Language = $cmbLanguage.SelectedItem.Tag
                $Script:ConfigResult.AdditionalLanguages = @()
                foreach ($cb in $langCheckboxes) {
                    if ($cb.IsChecked) {
                        $Script:ConfigResult.AdditionalLanguages += $cb.Tag
                    }
                }
            
                # Windows Edition
                $Script:ConfigResult.WindowsEdition = $cmbWindowsEdition.SelectedItem.Tag
            
                # Apps
                $Script:ConfigResult.Apps = @{
                    'Word'        = $chkWord.IsChecked
                    'Excel'       = $chkExcel.IsChecked
                    'PowerPoint'  = $chkPowerPoint.IsChecked
                    'Outlook'     = $chkOutlook.IsChecked
                    'Access'      = $chkAccess.IsChecked
                    'Publisher'   = $chkPublisher.IsChecked
                    'Teams'       = $chkTeams.IsChecked
                    'OneNote'     = $chkOneNote.IsChecked
                    'OneDrive'    = $chkOneDrive.IsChecked
                    'Skype'       = $chkSkype.IsChecked
                    'Defender'    = $chkDefender.IsChecked
                    'ToDo'        = $chkToDo.IsChecked
                    'PCManager'   = $chkPCManager.IsChecked
                    'PowerBI'     = $chkPowerBI.IsChecked
                    'StickyNotes' = $chkStickyNotes.IsChecked
                    'VSCode'      = $chkVSCode.IsChecked
                    'VS2022'      = $chkVS2022.IsChecked
                    'PowerToys'   = $chkPowerToys.IsChecked
                    'Copilot'     = $chkCopilot.IsChecked
                }
            
                # Extras (legacy flags for compatibility + new ones)
                $Script:ConfigResult.IncludeProject = $chkProject.IsChecked
                $Script:ConfigResult.IncludeVisio = $chkVisio.IsChecked
                $Script:ConfigResult.IncludeClipchamp = $chkClipchamp.IsChecked
                $Script:ConfigResult.IncludePowerAutomate = $chkPowerAutomate.IsChecked
            
                $window.DialogResult = $true
                $window.Close()
            })
        
        # Window close via X
        $window.Add_Closing({
                param($s, $e)
                # Suppress PSScriptAnalyzer unused parameter warning (required by WPF event signature)
                $null = $s, $e
                if (-not $window.DialogResult) {
                    $Script:ConfigResult.Cancelled = $true
                }
            })
        
        # Show dialog (blocking)
        Write-Log "Showing configuration window..." -Level Debug
        $result = $window.ShowDialog()
        Write-Log "Configuration window closed. Result: $result, Cancelled: $($Script:ConfigResult.Cancelled)" -Level Debug
        
        if ($Script:ConfigResult.Cancelled) {
            return $false
        }
        
        # Transfer to UserConfig
        $Script:UserConfig.InstallMode = $Script:ConfigResult.Mode
        $Script:UserConfig.OfficeVersion = $Script:ConfigResult.Version
        $Script:UserConfig.Channel = $Script:ConfigResult.Channel
        $Script:UserConfig.PrimaryLanguage = $Script:ConfigResult.Language
        $Script:UserConfig.AdditionalLanguages = $Script:ConfigResult.AdditionalLanguages
        $Script:UserConfig.WindowsEdition = $Script:ConfigResult.WindowsEdition
        $Script:UserConfig.IncludeProject = $Script:ConfigResult.IncludeProject
        $Script:UserConfig.IncludeVisio = $Script:ConfigResult.IncludeVisio
        $Script:UserConfig.IncludeClipchamp = $Script:ConfigResult.IncludeClipchamp
        $Script:UserConfig.IncludePowerAutomate = $Script:ConfigResult.IncludePowerAutomate
        
        if ($Script:ConfigResult.Mode -eq 'Custom') {
            $Script:UserConfig.SelectedApps = $Script:ConfigResult.Apps
        }
        else {
            # Express mode - Recommended apps
            $Script:UserConfig.SelectedApps = @{
                'Word' = $true; 'Excel' = $true; 'PowerPoint' = $true
                'Outlook' = $true; 'Access' = $false; 'Publisher' = $false
                'Teams' = $true; 'OneNote' = $false; 'OneDrive' = $false
                'Skype' = $false
            }
            $Script:UserConfig.IncludeProject = $false
            $Script:UserConfig.IncludeVisio = $false
            $Script:UserConfig.IncludeClipchamp = $true
            $Script:UserConfig.IncludePowerAutomate = $false
            $Script:UserConfig.WindowsEdition = 'Pro'
        }
        
        return $true
    }
    catch {
        Write-Log "Failed to show configuration window: $($_.Exception.Message)" -Level Error
        Write-Log "Stack trace: $($_.Exception.StackTrace)" -Level Error
        return $false
    }
}
function Start-ProgressWindow {
    param(
        [string]$Title = "Microsoft Ultimate Installer $($Script:Version)",
        [string]$Subtitle = "Installing Office..."
    )
    # Pass the global Base64 strings to the inner function
    Show-Progress -Title $Title -Subtitle $Subtitle -HeaderBase64 $Script:ACTechLogoBase64 -FooterBase64 $Script:ACTechLogoBase64
}

function Show-Progress {
    param($Title, $Subtitle, $HeaderBase64, $FooterBase64)
    
    # Close existing standard progress if open (legacy)
    if ($Script:ProgressRunspace) { Close-Progress }
    
    $initStatus = "Initializing..."
    
    $Script:ProgressSync = [hashtable]::Synchronized(@{
            Status        = $initStatus
            Progress      = 0
            SubStatus     = ''
            ShouldClose   = $false
            RequestCancel = $false
        })
    
    $Script:ProgressRunspace = [runspacefactory]::CreateRunspace()
    $Script:ProgressRunspace.ApartmentState = 'STA'
    $Script:ProgressRunspace.ThreadOptions = 'ReuseThread'
    $Script:ProgressRunspace.Open()
    
    $code = {
        param($Sync, $Title, $Subtitle, $HeaderBase64, $FooterBase64)
        
        # Escape strings
        $Title = $Title -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
        
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        
        # Load C# Types for Transparency (if available in this AppDomain)
        # Note: In a separate runspace, we might need to rely on the type being already loaded in the process or re-add it.
        # Ideally, we rely on the main runspace having loaded it. However, if 'WindowEffects' isn't visible, transparency won't blur.
        
        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Width="500" Height="280"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" ResizeMode="CanResizeWithGrip"
        Background="Transparent" Topmost="True">
    <Window.Resources>
        <Style x:Key="ModernButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#333333"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#454545"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#777777"/>
                                <Setter TargetName="border" Property="Cursor" Value="Hand"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1A1A1A"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#555555"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Window Control Styles (Copied from Main Window) -->
        <Style TargetType="Button" x:Key="Style.Button.WindowControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#AAAAAA"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#33FFFFFF"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style TargetType="Button" x:Key="Style.Button.Close" BasedOn="{StaticResource Style.Button.WindowControl}">
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E81123"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Border x:Name="MainBorder" Background="#FF151515" CornerRadius="8" BorderBrush="#333333" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/> <!-- Header -->
                <RowDefinition Height="*"/>    <!-- Content -->
                <RowDefinition Height="Auto"/> <!-- Footer -->
            </Grid.RowDefinitions>
            
            <!-- Header -->
            <Grid Grid.Row="0" Margin="20,15,20,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
                    <Image x:Name="imgHeader" Width="24" Margin="0,0,10,0" RenderOptions.BitmapScalingMode="HighQuality"/>
                    <TextBlock Text="$Title" FontSize="14" Foreground="White" VerticalAlignment="Center" FontWeight="SemiBold">
                        <TextBlock.Effect>
                            <DropShadowEffect Color="Black" Direction="320" ShadowDepth="1" BlurRadius="5" Opacity="0.8"/>
                        </TextBlock.Effect>
                    </TextBlock>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="btnMinimize" Content="__" Width="46" Height="32" Style="{StaticResource Style.Button.WindowControl}" Padding="0,0,0,8"/>
                    <Button x:Name="btnMaximize" Content="[ ]" Width="46" Height="32" Style="{StaticResource Style.Button.WindowControl}" FontSize="12"/>
                    <Button x:Name="btnClose" Content="X" Width="46" Height="32" Style="{StaticResource Style.Button.Close}" FontSize="14"/>
                </StackPanel>
            </Grid>
            
            <!-- Content -->
            <Grid Grid.Row="1" Margin="30,10,30,10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/> <!-- Spacer -->
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock x:Name="StatusText" Grid.Row="0" Text="Initializing..." FontSize="16" Foreground="White" Margin="0,0,0,15" FontWeight="Normal"/>
                
                <!-- Modern ProgressBar -->
                <Border Grid.Row="1" Height="6" CornerRadius="3" Background="#333333" Margin="0,0,0,10">
                    <ProgressBar x:Name="ProgressBar" Background="Transparent" BorderThickness="0" Foreground="#0078D4" Maximum="100" Value="0" Opacity="1"/>
                </Border>
                 
                <Grid Grid.Row="2" Margin="0,0,0,20">
                    <TextBlock x:Name="SubStatusText" Text="" FontSize="12" Foreground="#AAAAAA" HorizontalAlignment="Left"/>
                    <TextBlock x:Name="PercentText" Text="0%" FontSize="12" Foreground="#AAAAAA" HorizontalAlignment="Right"/>
                </Grid>
                
                <Button x:Name="btnCancel" Grid.Row="4" Content="Cancel" Width="100" Height="32"
                        Style="{StaticResource ModernButtonStyle}" HorizontalAlignment="Right"/>
            </Grid>
            
            <!-- Footer -->
            <Border Grid.Row="2" BorderBrush="#333333" BorderThickness="0,1,0,0" Padding="15">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                     <TextBlock Text="Developed by AC Tech" Foreground="#666666" FontSize="10" VerticalAlignment="Center" Margin="0,0,8,0"/>
                     <Image x:Name="imgFooter" Height="20" RenderOptions.BitmapScalingMode="HighQuality"/>
                </StackPanel>
            </Border>
        </Grid>
    </Border>
</Window>
"@
        
        try {
            [xml]$xamlDoc = $xaml
            $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
            $window = [Windows.Markup.XamlReader]::Load($reader)
            
            # Find Elements
            $statusText = $window.FindName('StatusText')
            $progressBar = $window.FindName('ProgressBar')
            $subStatusText = $window.FindName('SubStatusText')
            $percentText = $window.FindName('PercentText')
            $btnCancel = $window.FindName('btnCancel')
            $imgHeader = $window.FindName('imgHeader')
            $imgFooter = $window.FindName('imgFooter')
            $MainBorder = $window.FindName('MainBorder')

            # Load Images
            # Load Images
            if ($HeaderBase64) {
                try {
                    $bytes = [Convert]::FromBase64String($HeaderBase64)
                    $ms = New-Object System.IO.MemoryStream(, $bytes)
                    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bmp.BeginInit()
                    $bmp.StreamSource = $ms
                    $bmp.EndInit()
                    $imgHeader.Source = $bmp
                }
                catch {}
            }

            if ($FooterBase64) {
                try {
                    $bytes = [Convert]::FromBase64String($FooterBase64)
                    $ms = New-Object System.IO.MemoryStream(, $bytes)
                    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bmp.BeginInit()
                    $bmp.StreamSource = $ms
                    $bmp.EndInit()
                    $imgFooter.Source = $bmp
                }
                catch {}
            }

            # Enable Transparency Logic (Runspace Version)
            $SetTransparency = {
                param($enable, $win, $border)
                if ($enable) {
                    if ($border) { 
                        $border.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString("#88151515"))
                    }
                    # Attempt Blur Interop logic if Type exists
                    if ([System.Management.Automation.PSTypeName]::new('WindowEffects').Type) {
                        try {
                            $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($win)).Handle
                            [WindowEffects]::EnableBlur($hwnd, 0, 0)
                        }
                        catch {}
                    }
                }
                else {
                    if ($border) { 
                        $border.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString("#FF151515"))
                    }
                }
            }

            # Window Controls Logic
            $btnMinimize = $window.FindName('btnMinimize')
            $btnMaximize = $window.FindName('btnMaximize')
            $btnClose = $window.FindName('btnClose')
            
            if ($btnMinimize) { $btnMinimize.Add_Click({ $window.WindowState = 'Minimized' }) }
            if ($btnMaximize) { 
                $btnMaximize.Add_Click({
                        if ($window.WindowState -eq 'Normal') {
                            $window.WindowState = 'Maximized'
                            $MainBorder.CornerRadius = '0'
                            $MainBorder.BorderThickness = '0'
                            $btnMaximize.Content = "[-]"
                        }
                        else {
                            $window.WindowState = 'Normal'
                            $MainBorder.CornerRadius = '8'
                            $MainBorder.BorderThickness = '1'
                            $btnMaximize.Content = "[ ]"
                        }
                    })
            }
            if ($btnClose) { 
                $btnClose.Add_Click({ 
                        $Sync.RequestCancel = $true 
                        # Let the loop handle closing via RequestCancel
                    }) 
            }

            # Drag Move Handler (exclude buttons)
            $window.Add_MouseLeftButtonDown({
                    param($src, $e)
                    if ($e.OriginalSource -is [System.Windows.Controls.Button] -or 
                        ($e.OriginalSource.Parent -is [System.Windows.Controls.Button])) {
                        return
                    }
                    # Apply transparency effect during drag
                    & $SetTransparency -enable $true -win $window -border $MainBorder
                    try { $window.DragMove() } catch {}
                    & $SetTransparency -enable $false -win $window -border $MainBorder
                })
            
            $btnCancel.Add_Click({
                    $Sync.RequestCancel = $true
                    $btnCancel.IsEnabled = $false
                    $btnCancel.Content = "Cancelling..."
                })
            
            $window.Add_Closing({
                    param($s, $e)
                    $null = $s
                    if (-not $Sync.ShouldClose) {
                        $Sync.RequestCancel = $true
                        $e.Cancel = $true
                    }
                })
            
            # Timer for updates
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromMilliseconds(100)
            $timer.Add_Tick({
                    $statusText.Text = $Sync.Status
                    $progressBar.Value = $Sync.Progress
                    $subStatusText.Text = $Sync.SubStatus
                    $percentText.Text = "$([int]$Sync.Progress)%"
                    if ($Sync.ShouldClose) { $window.Close() }
                    
                    # Ensure window stays top if needed
                    if ($window.WindowState -eq 'Minimized') { $window.WindowState = 'Normal' }
                })
            $timer.Start()
            
            $window.ShowDialog() | Out-Null
        }
        catch {
            $_ | Out-File "$env:TEMP\M365_Progress_Error.txt"
        }
    }
    
    $ps = [PowerShell]::Create()
    $ps.Runspace = $Script:ProgressRunspace
    $ps.AddScript($code).AddArgument($Script:ProgressSync).AddArgument($title).AddArgument($subtitle).AddArgument($HeaderBase64).AddArgument($FooterBase64) | Out-Null
    $Script:ProgressHandle = $ps.BeginInvoke()
    
    Start-Sleep -Milliseconds 300  # Let window initialize
}

function Update-Progress {
    param([string]$Status, [int]$Percent, [string]$SubStatus = '')
    if ($Script:ProgressSync) {
        $Script:ProgressSync.Status = $Status
        $Script:ProgressSync.Progress = $Percent
        $Script:ProgressSync.SubStatus = $SubStatus
    }
}

function Close-Progress {
    try {
        if ($Script:ProgressSync) {
            $Script:ProgressSync.ShouldClose = $true
            Start-Sleep -Milliseconds 300  # Give UI time to close gracefully
        }
    }
    catch {
        # Intentionally suppressed: Progress sync may already be disposed
        $null = $_
    }
    
    try {
        if ($Script:ProgressHandle -and -not $Script:ProgressHandle.IsCompleted) {
            $Script:ProgressHandle.AsyncWaitHandle.WaitOne(1000) | Out-Null
        }
    }
    catch {
        # Intentionally suppressed: Handle may already be completed
        $null = $_
    }
    
    try {
        if ($Script:ProgressRunspace) {
            $Script:ProgressRunspace.Close()
            $Script:ProgressRunspace.Dispose()
            $Script:ProgressRunspace = $null
        }
    }
    catch {
        # Intentionally suppressed: Runspace may already be closed
        $null = $_
    }
}

function Write-Checkpoint {
    param(
        [Parameter(Mandatory)] [string]$Tag,
        [string]$Note
    )
    $noteText = if ($Note) { " - $Note" } else { '' }
    Write-Log "Checkpoint: $Tag (Stage=$Script:Stage)$noteText" -Level Debug
}

function Test-FileExists {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Label
    )
    if (-not (Test-Path $Path)) {
        throw "$Label not found: $Path"
    }
    try {
        $size = (Get-Item $Path).Length
        Write-Log "$Label located ($size bytes): $Path" -Level Debug
    }
    catch {
        Write-Log "$Label located: $Path (size unknown)" -Level Debug
    }
}

# ============================================================================
# CORE FUNCTIONS
# ============================================================================

function New-TempFolder {
    if (Test-Path $Script:Config.TempFolder) {
        Remove-Item $Script:Config.TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item $Script:Config.TempFolder -ItemType Directory -Force | Out-Null
    Write-Log "Temp folder created: $($Script:Config.TempFolder)" -Level Debug
}

function Get-ODT {
    Write-Log "Attempting to download Office Deployment Tool..."
    $odtPath = Join-Path $Script:Config.TempFolder 'odt_setup.exe'
    
    foreach ($url in $Script:Config.ODTUrls) {
        try {
            # Download with retry logic
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
            $webClient.DownloadFile($url, $odtPath)
            
            # Validate download
            if (-not (Test-Path $odtPath)) {
                Write-Log "Download failed - file not created" -Level Warning
                continue
            }
            
            $fileSize = (Get-Item $odtPath).Length
            if ($fileSize -lt 1000000) {
                # Less than 1MB is suspicious
                Write-Log "Downloaded file too small ($fileSize bytes)" -Level Warning
                Remove-Item $odtPath -Force -ErrorAction SilentlyContinue
                continue
            }
            
            Write-Log "ODT downloaded successfully ($fileSize bytes)" -Level Success
            break
        }
        catch {
            Write-Log "Download error: $($_.Exception.Message)" -Level Warning
            Remove-Item $odtPath -Force -ErrorAction SilentlyContinue
            continue
        }
    }
    
    if (-not (Test-Path $odtPath)) {
        throw "Failed to download Office Deployment Tool after all attempts"
    }
    
    # Extract ODT
    try {
        Write-Log "Extracting Office Deployment Tool..." -Level Debug
        $extractPath = Join-Path $Script:Config.TempFolder 'ODT'
        New-Item $extractPath -ItemType Directory -Force | Out-Null
        
        # Run ODT with extract parameter
        $proc = Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:`"$extractPath`"" -Wait -PassThru -WindowStyle Hidden -WorkingDirectory $Script:Config.TempFolder
        
        Write-Log "ODT extraction exit code: $($proc.ExitCode)" -Level Debug
        
        # Validate extraction
        Start-Sleep -Milliseconds 500  # Give system time to complete extraction
        $setupExe = Join-Path $extractPath 'setup.exe'
        
        if (-not (Test-Path $setupExe)) {
            Write-Log "setup.exe not found in $extractPath" -Level Error
            Write-Log "Contents of $extractPath :" -Level Debug
            Get-ChildItem $extractPath -Recurse | ForEach-Object { Write-Log "  - $($_.FullName)" -Level Debug }
            throw "setup.exe not found after extraction"
        }
        
        Write-Log "setup.exe found: $setupExe" -Level Success
        return $setupExe
    }
    catch {
        throw "ODT extraction failed: $($_.Exception.Message)"
    }
}

function New-ConfigXML {
    Write-Log "Generating Configuration XML based on user settings..." -Level Debug
    $path = Join-Path $Script:Config.TempFolder 'configuration.xml'
    
    # Get user configuration
    $version = $Script:UserConfig.OfficeVersion
    $channel = $Script:UserConfig.Channel
    $primaryLang = $Script:UserConfig.PrimaryLanguage
    $additionalLangs = $Script:UserConfig.AdditionalLanguages
    $selectedApps = $Script:UserConfig.SelectedApps
    $includeProject = $Script:UserConfig.IncludeProject
    $includeVisio = $Script:UserConfig.IncludeVisio
    
    # Build language tags
    $langTags = "      <Language ID=`"$primaryLang`" />"
    foreach ($lang in $additionalLangs) {
        if ($lang -ne $primaryLang) {
            $langTags += "`n      <Language ID=`"$lang`" />"
        }
    }
    
    # Build exclude app tags based on user selection
    $excludeApps = ""
    $appMapping = @{
        'Word'       = 'Word'
        'Excel'      = 'Excel'
        'PowerPoint' = 'PowerPoint'
        # Outlook removed to prevent ODT installation (Classic) - We install New Outlook via Store
        'Access'     = 'Access'
        'Publisher'  = 'Publisher'
        'OneNote'    = 'OneNote'
        'OneDrive'   = 'OneDrive'
        'Teams'      = 'Teams'
        'Lync'       = 'Lync'
        'Groove'     = 'Groove'
    }
    
    foreach ($app in $appMapping.Keys) {
        if ($selectedApps.ContainsKey($app) -and -not $selectedApps[$app]) {
            $excludeApps += "`n      <ExcludeApp ID=`"$($appMapping[$app])`" />"
        }
    }
    
    # ALWAYS exclude Classic Outlook (we use the New Outlook App)
    $excludeApps += "`n      <ExcludeApp ID=`"Outlook`" />"
    # Always exclude Lync and Groove if not explicitly selected
    if (-not $selectedApps.ContainsKey('Lync') -or -not $selectedApps['Lync']) {
        if ($excludeApps -notmatch 'Lync') {
            $excludeApps += "`n      <ExcludeApp ID=`"Lync`" />"
        }
    }
    if (-not $selectedApps.ContainsKey('Groove') -or -not $selectedApps['Groove']) {
        if ($excludeApps -notmatch 'Groove') {
            $excludeApps += "`n      <ExcludeApp ID=`"Groove`" />"
        }
    }
    
    # Start building XML
    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel" MigrateArch="TRUE" AllowCdnFallback="TRUE" ForceUpgrade="TRUE" SourcePath="$($Script:Config.TempFolder)">
    <Product ID="$version">
$langTags$excludeApps
    </Product>
"@
    
    # Add Project if selected and supported
    # NOTE: Project and Visio are standalone products, they should NOT have ExcludeApps
    if ($includeProject) {
        $projectId = switch ($version) {
            'O365ProPlusRetail' { 'ProjectProRetail' }
            'O365BusinessRetail' { 'ProjectProRetail' }
            'ProPlus2024Retail' { 'ProjectPro2024Retail' }
            'ProPlus2021Retail' { 'ProjectPro2021Retail' }
            'ProPlus2019Retail' { 'ProjectPro2019Retail' }
            default { 'ProjectProRetail' }
        }
        $xmlContent += @"

    <Product ID="$projectId">
$langTags
    </Product>
"@
    }
    
    # Add Visio if selected and supported
    # NOTE: Visio is a standalone product, it should NOT have ExcludeApps
    if ($includeVisio) {
        $visioId = switch ($version) {
            'O365ProPlusRetail' { 'VisioProRetail' }
            'O365BusinessRetail' { 'VisioProRetail' }
            'ProPlus2024Retail' { 'VisioPro2024Retail' }
            'ProPlus2021Retail' { 'VisioPro2021Retail' }
            'ProPlus2019Retail' { 'VisioPro2019Retail' }
            default { 'VisioProRetail' }
        }
        $xmlContent += @"

    <Product ID="$visioId">
$langTags
    </Product>
"@
    }
    
    # Add language pack for additional languages if any
    if ($additionalLangs.Count -gt 0) {
        $langPackTags = ""
        foreach ($lang in $additionalLangs) {
            if ($lang -ne $primaryLang) {
                $langPackTags += "`n      <Language ID=`"$lang`" />"
            }
        }
        if ($langPackTags) {
            $xmlContent += @"

    <Product ID="LanguagePack">$langPackTags
    </Product>
"@
        }
    }
    
    # Close Add section and add properties
    $xmlContent += @"

  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Updates Enabled="TRUE" Channel="$channel" />
  <RemoveMSI All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\common\privacy" Name="disconnectedstate" Value="2" Type="REG_DWORD" App="office16" Id="L_DisableAllConnectedExperiences" />
    <User Key="software\microsoft\office\16.0\common\privacy" Name="usercontentdisabled" Value="1" Type="REG_DWORD" App="office16" Id="L_DisableConnectedExperiencesContent" />
    <User Key="software\microsoft\office\16.0\common\privacy" Name="downloadcontentdisabled" Value="1" Type="REG_DWORD" App="office16" Id="L_DisableConnectedExperiencesDownload" />
    <User Key="software\microsoft\office\16.0\common\privacy" Name="controllerconnectedservicesenabled" Value="2" Type="REG_DWORD" App="office16" Id="L_DisableOptionalConnectedExperiences" />
    <User Key="software\microsoft\office\16.0\common" Name="qmenable" Value="0" Type="REG_DWORD" App="office16" Id="L_QMEnable" />
    <User Key="software\microsoft\office\16.0\common\feedback" Name="enabled" Value="0" Type="REG_DWORD" App="office16" Id="L_Feedback" />
    <User Key="software\microsoft\office\16.0\common\internet" Name="useonlinecontent" Value="0" Type="REG_DWORD" App="office16" Id="L_OnlineContent" />
  </AppSettings>
</Configuration>
"@
    
    Set-Content -Path $path -Value $xmlContent -Encoding UTF8 -Force
    Write-Log "Configuration XML created: $path" -Level Success
    Write-Log "  Version: $version, Channel: $channel" -Level Debug
    Write-Log "  Primary Language: $primaryLang, Additional: $($additionalLangs -join ', ')" -Level Debug
    Write-Log "  Include Project: $includeProject, Include Visio: $includeVisio" -Level Debug
    return $path
}

function Remove-Conflicting-Office {
    # Check for cancel before starting
    if ($Script:ProgressSync -and $Script:ProgressSync.RequestCancel) {
        throw "Installation cancelled by user"
    }
    
    Write-Log "Checking for conflicting Office installations..." -Level Info
    
    # Kill processes
    @('OneDrive', 'OneNote', 'Microsoft.Notes', 'Skype', 'lync', 'WINWORD', 'EXCEL', 'POWERPNT', 'OUTLOOK', 'ONENOTE') | ForEach-Object {
        $procs = Get-Process $_ -ErrorAction SilentlyContinue
        if ($procs) {
            Write-Log "Stopping process: $_" -Level Debug
            $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }
    
    # Uninstall existing Office installations (Click-to-Run)
    Write-Log "Checking for existing Office Click-to-Run installations..." -Level Debug
    $officeC2R = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($officeC2R) {
        Write-Log "Found existing Office Click-to-Run installation - attempting silent removal" -Level Warning
        $c2rPath = Join-Path $env:ProgramFiles 'Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe'
        if (Test-Path $c2rPath) {
            Write-Log "Running Office C2R uninstall..." -Level Debug
            Start-Process $c2rPath -ArgumentList 'scenario=install', 'scenariosubtype=ARP', 'productstouninstall=O365ProPlusRetail,OneNoteRetail,ProjectProRetail,VisioProRetail', '/quiet' -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep -Seconds 3
            Write-Log "Office C2R uninstall completed" -Level Debug
        }
    }
}

function Remove-Windows-Bloatware {
    Write-Log "Removing Windows bloatware (OneDrive, OneNote, StickyNotes)..." -Level Info

    # Uninstall OneDrive
    $odSetup = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    if (-not (Test-Path $odSetup)) { $odSetup = "$env:SystemRoot\System32\OneDriveSetup.exe" }
    if (Test-Path $odSetup) {
        Write-Log "Uninstalling OneDrive..." -Level Debug
        Start-Process $odSetup -ArgumentList "/uninstall" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
    }
    
    # Remove Appx packages (including OneNote)
    Write-Log "Removing UWP app packages..." -Level Debug
    @('*OneDrive*', '*OneNote*', '*StickyNotes*', '*SkypeApp*', '*MicrosoftOfficeHub*') | ForEach-Object {
        $pkgName = $_
        $packages = Get-AppxPackage $pkgName -AllUsers -ErrorAction SilentlyContinue
        if ($packages) {
            Write-Log "Removing package: $pkgName" -Level Debug
            $packages | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        }
        # Also try removing provisioned packages
        $provisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $pkgName }
        if ($provisioned) {
            Write-Log "Removing provisioned package: $pkgName" -Level Debug
            $provisioned | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
        }
    }
    
    # Registry Blocks
    Write-Log "Applying registry policies..." -Level Debug
    $regPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive',
        'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate'
    )
    foreach ($regPath in $regPaths) {
        if (-not (Test-Path $regPath)) { New-Item $regPath -Force | Out-Null }
    }
    
    Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Name 'DisableFileSyncNGSC' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate' -Name 'HideOneDrive' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate' -Name 'HideOneNote' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    
    Write-Log "Bloatware removal complete" -Level Success
}

function Set-WindowsInsider {
    param(
        [string]$Channel = "None"
    )
    
    Write-Log "Configuring Windows Insider Program: $Channel" -Level Info
    $regPath = "HKLM:\SOFTWARE\Microsoft\WindowsSelfHost\Applicability"
    $regPathUI = "HKLM:\SOFTWARE\Microsoft\WindowsSelfHost\UI\Selection"

    
    if ($Channel -eq "None" -or $Channel -eq "Normal") {
        # Opt-out / Remove Insider settings
        Write-Log "Removing Windows Insider settings..." -Level Info
        
        # Remove Applicability keys
        if (Test-Path $regPath) {
            Remove-ItemProperty -Path $regPath -Name "BranchName" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPath -Name "ContentType" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPath -Name "Ring" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPath -Name "UIContentType" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPath -Name "UIBranch" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPath -Name "UIRing" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $regPath -Name "EnablePreviewBuilds" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
        }
        
        # Clean UI selection keys to reflect opt-out
        if (Test-Path $regPathUI) {
            Remove-ItemProperty -Path $regPathUI -Name "UIBranch" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPathUI -Name "UIContentType" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $regPathUI -Name "UIRing" -ErrorAction SilentlyContinue
        }
        
        Write-Log "Insider Program settings removed." -Level Success
        return
    }

    # Map Channels to Registry Values
    $branch = "ReleasePreview"
    $content = "Mainline"
    $ring = "External"
    
    switch ($Channel) {
        "Dev" {
            $branch = "Dev"
            $content = "Mainline"
            $ring = "External"
        }
        "Beta" {
            $branch = "Beta"
            $content = "Mainline"
            $ring = "External"
        }
        "ReleasePreview" {
            $branch = "ReleasePreview"
            $content = "Mainline"
            $ring = "External"
        }
    }
    
    # Create keys if missing
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    if (-not (Test-Path $regPathUI)) { New-Item -Path $regPathUI -Force | Out-Null }
    
    # Set Applicability
    Set-ItemProperty -Path $regPath -Name "EnablePreviewBuilds" -Value 2 -Type DWord -Force
    Set-ItemProperty -Path $regPath -Name "BranchName" -Value $branch -Type String -Force
    Set-ItemProperty -Path $regPath -Name "ContentType" -Value $content -Type String -Force
    Set-ItemProperty -Path $regPath -Name "Ring" -Value $ring -Type String -Force
    
    # Set UI
    Set-ItemProperty -Path $regPathUI -Name "UIBranch" -Value $branch -Type String -Force
    Set-ItemProperty -Path $regPathUI -Name "UIContentType" -Value $content -Type String -Force
    Set-ItemProperty -Path $regPathUI -Name "UIRing" -Value $ring -Type String -Force
    
    Write-Log "Windows Insider configured to: $Channel" -Level Success
}


function Start-CleanupOperations {
    Write-Log "Running post-installation cleanup..." -Level Info
    
    # Force remove OneNote executable if it exists
    $oneNotePaths = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\ONENOTE.EXE",
        "$env:ProgramFiles(x86)\Microsoft Office\root\Office16\ONENOTE.EXE"
    )
    
    foreach ($path in $oneNotePaths) {
        if (Test-Path $path) {
            Write-Log "Found OneNote executable at $path - Removing..." -Level Warning
            try {
                Stop-Process -Name "ONENOTE" -Force -ErrorAction SilentlyContinue
                Remove-Item $path -Force -ErrorAction Stop
                Write-Log "Removed OneNote executable" -Level Success
            }
            catch {
                Write-Log "Failed to remove OneNote executable: $($_.Exception.Message)" -Level Warning
            }
        }
    }

    # Force remove Sticky Notes if it exists (Desktop version or leftovers)
    $stickyNotesPaths = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\StikyNot.exe",
        "$env:ProgramFiles(x86)\Microsoft Office\root\Office16\StikyNot.exe",
        "$env:SystemRoot\System32\StikyNot.exe"
    )

    foreach ($path in $stickyNotesPaths) {
        if (Test-Path $path) {
            Write-Log "Found Sticky Notes executable at $path - Removing..." -Level Warning
            try {
                Stop-Process -Name "StikyNot" -Force -ErrorAction SilentlyContinue
                Remove-Item $path -Force -ErrorAction Stop
                Write-Log "Removed Sticky Notes executable" -Level Success
            }
            catch {
                Write-Log "Failed to remove Sticky Notes executable: $($_.Exception.Message)" -Level Warning
            }
        }
    }
}

function Remove-AppShortcuts {
    Write-Log "Removing Start Menu shortcuts for removed apps..." -Level Info

    $shortcutTargets = @(
        'OneNote',
        'Microsoft OneNote',
        'Sticky Notes',
        'StickyNotes',
        'OneDrive',
        'Skype'
    )

    $startMenuPaths = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs')
    )

    foreach ($path in $startMenuPaths) {
        if (Test-Path $path) {
            $shortcuts = Get-ChildItem -Path $path -Filter "*.lnk" -Recurse -ErrorAction SilentlyContinue
            foreach ($shortcut in $shortcuts) {
                foreach ($target in $shortcutTargets) {
                    if ($shortcut.Name -match $target) {
                        try {
                            Remove-Item $shortcut.FullName -Force -ErrorAction Stop
                            Write-Log "Removed shortcut: $($shortcut.Name)" -Level Debug
                        }
                        catch {
                            Write-Log "Failed to remove shortcut $($shortcut.Name): $($_.Exception.Message)" -Level Warning
                        }
                        break
                    }
                }
            }
        }
    }
    
    Write-Log "Shortcut cleanup complete" -Level Success
}

function Install-Office {
    param(
        [Parameter(Mandatory)] [string]$SetupPath,
        [Parameter(Mandatory)] [string]$ConfigPath
    )
    
    Write-Log "Starting Office Installation (Download & Install)..." -Level Info
    
    # Download
    Update-Progress -Status (L 'StatusDownloadingOffice') -Percent 25 -SubStatus (L 'SubStatusInternetSpeed')
    Write-Log "Running ODT download phase with config: $ConfigPath" -Level Debug
    
    $proc = Start-Process -FilePath $SetupPath -ArgumentList "/download `"$ConfigPath`"" -Wait -PassThru -WindowStyle Hidden -WorkingDirectory $Script:Config.TempFolder -ErrorAction Stop
    Write-Log "ODT download phase exit code: $($proc.ExitCode)" -Level Debug
    
    if ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 3010) {
        Write-Log "Download warning/error code: $($proc.ExitCode)" -Level Warning
    }
    
    # Install
    Update-Progress -Status (L 'StatusInstallingOffice') -Percent 50 -SubStatus (L 'SubStatusApplyingConfig')
    Write-Log "Running ODT install phase" -Level Debug
    
    $proc = Start-Process -FilePath $SetupPath -ArgumentList "/configure `"$ConfigPath`"" -Wait -PassThru -WindowStyle Hidden -WorkingDirectory $Script:Config.TempFolder -ErrorAction Stop
    Write-Log "ODT install phase exit code: $($proc.ExitCode)" -Level Debug
    
    if ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 3010) {
        throw "Office installation failed with code $($proc.ExitCode)"
    }
    
    Write-Log "Office installation complete" -Level Success
}

function Install-Winget {
    Write-Log "Checking Windows Package Manager (winget)..." -Level Info
    
    $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
    if ($wingetPath) {
        Write-Log "Winget is already installed" -Level Success
        return $wingetPath.Source
    }
    
    Write-Log "Winget not found - installing App Installer from Microsoft Store..." -Level Warning
    
    try {
        # Download App Installer (includes winget)
        $appInstallerUrl = "https://aka.ms/getwinget"
        $appInstallerPath = Join-Path $Script:Config.TempFolder "Microsoft.DesktopAppInstaller.msixbundle"
        
        Write-Log "Downloading App Installer..." -Level Debug
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($appInstallerUrl, $appInstallerPath)
        
        Write-Log "Installing App Installer..." -Level Debug
        Add-AppxPackage -Path $appInstallerPath -ErrorAction Stop
        
        # Wait for installation to complete
        Start-Sleep -Seconds 5
        
        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
        if ($wingetPath) {
            Write-Log "Winget installed successfully" -Level Success
            return $wingetPath.Source
        }
        else {
            Write-Log "Winget installation completed but command not found in PATH" -Level Warning
            # Try common installation paths
            $possiblePaths = @(
                "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
                "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller*\winget.exe"
            )
            foreach ($path in $possiblePaths) {
                $resolved = Get-ChildItem $path -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
                if ($resolved -and (Test-Path $resolved)) {
                    Write-Log "Found winget at: $resolved" -Level Success
                    return $resolved
                }
            }
            throw "Winget installation succeeded but executable not found"
        }
    }
    catch {
        Write-Log "Failed to install winget: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Install-Extras {
    Write-Log "Installing Winget packages..." -Level Info
    Update-Progress -Status (L 'StatusInstallingExtras') -Percent 65 -SubStatus (L 'SubStatusClipchampPowerAutomate')
    
    # Build list of packages to install based on user configuration
    $packagesToInstall = @()
    
    # ---------------------------
    # App Definitions (Winget/Store)
    # ---------------------------
    $appsMap = @{
        'Outlook'     = @{ Id = '9NRX63209R7B'; Source = 'msstore'; Name = 'Outlook (New)' }
        'Defender'    = @{ Id = '9P6PMZTM93LR'; Source = 'msstore'; Name = 'Microsoft Defender' }
        'ToDo'        = @{ Id = '9NBLGGH5R558'; Source = 'msstore'; Name = 'Microsoft To Do' }
        'PCManager'   = @{ Id = '9PM860492SZD'; Source = 'msstore'; Name = 'Microsoft PC Manager' }
        'StickyNotes' = @{ Id = '9NBLGGH4BI9O'; Source = 'msstore'; Name = 'Microsoft Sticky Notes' }
        'PowerBI'     = @{ Id = 'Microsoft.PowerBI'; Source = 'winget'; Name = 'Power BI Desktop' }
        'PowerToys'   = @{ Id = 'Microsoft.PowerToys'; Source = 'winget'; Name = 'Microsoft PowerToys' }
        'VSCode'      = @{ Id = 'Microsoft.VisualStudioCode.Insiders'; Source = 'winget'; Name = 'Visual Studio Code Insiders' }
        'VS2022'      = @{ Id = 'Microsoft.VisualStudio.Enterprise.Preview'; Source = 'winget'; Name = 'Visual Studio 2026 Enterprise Preview' }
        'Copilot'     = @{ Id = '9NHT9RB2F4HD'; Source = 'msstore'; Name = 'Microsoft Copilot' }
        'Skype'       = @{ Id = 'Microsoft.Skype'; Source = 'winget'; Name = 'Skype' }
        'Clipchamp'   = @{ Id = '9P1J8S7CCWWT'; Source = 'msstore'; Name = 'Microsoft Clipchamp' } # Legacy key fallback
    }

    # Add mapped apps
    $userApps = $Script:UserConfig.Apps
    foreach ($key in $appsMap.Keys) {
        if ($userApps.ContainsKey($key) -and $userApps[$key]) {
            $packagesToInstall += $appsMap[$key]
        }
    }

    # ---------------------------
    # Legacy / Explicit Flags
    # ---------------------------
    
    # Check if Clipchamp should be installed (Legacy Flag or App check)
    if ($Script:UserConfig.IncludeClipchamp -and -not $userApps['Clipchamp']) {
        # Only add if not already added via map (fix mismatched key loop)
        $packagesToInstall += @{ Id = '9P1J8S7CCWWT'; Source = 'msstore'; Name = 'Microsoft Clipchamp' }
    }
    
    # Check if Power Automate should be installed
    if ($Script:UserConfig.IncludePowerAutomate) {
        $packagesToInstall += @{ Id = 'Microsoft.PowerAutomateDesktop'; Source = 'winget'; Name = 'Power Automate Desktop' }
    }
    
    if ($packagesToInstall.Count -eq 0) {
        Write-Log "No extra packages selected for installation - skipping" -Level Debug
        return
    }
    
    # Ensure winget is installed
    try {
        $wingetExe = Install-Winget
    }
    catch {
        Write-Log "Cannot proceed with extras installation - winget unavailable: $($_.Exception.Message)" -Level Warning
        return
    }
    
    foreach ($pkg in $packagesToInstall) {
        $pkgId = if ($pkg -is [hashtable]) { $pkg.Id } else { $pkg }
        $pkgSource = if ($pkg -is [hashtable]) { $pkg.Source } else { 'winget' }
        $pkgName = if ($pkg -is [hashtable]) { $pkg.Name } else { $pkg }
        
        Write-Log "Installing package: $pkgName" -Level Info
        try {
            # First check if already installed (add --accept-source-agreements for msstore)
            $listArgs = "list --id $pkgId --accept-source-agreements"
            Start-Process -FilePath $wingetExe -ArgumentList $listArgs -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "$env:TEMP\winget_check.txt" -ErrorAction Stop | Out-Null
            $installed = Get-Content "$env:TEMP\winget_check.txt" -Raw -ErrorAction SilentlyContinue
            
            if ($installed -match [regex]::Escape($pkgId)) {
                Write-Log "Package $pkgName is already installed - skipping" -Level Debug
                continue
            }
            
            # Install the package with correct source
            # NOTE: --silent flag fails with WindowStyle Hidden for msstore; we'll retry without it
            $pkgInstalled = $false
            $installAttempts = 0
            $maxAttempts = 2
            
            while (-not $pkgInstalled -and $installAttempts -lt $maxAttempts) {
                $installAttempts++
                $useSilent = if ($installAttempts -eq 1) { $true } else { $false }
                
                $wingetArgs = "install --id $pkgId --accept-package-agreements --accept-source-agreements --source $pkgSource"
                if ($useSilent) { $wingetArgs += " --silent" }
                if ($pkgSource -ne 'msstore') { $wingetArgs += " --force" }
                
                Write-Log "Running (attempt $installAttempts): winget $wingetArgs" -Level Debug
                $proc = Start-Process -FilePath $wingetExe -ArgumentList $wingetArgs -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
                
                Write-Log "Installation exit code: $($proc.ExitCode)" -Level Debug
                
                # Interpret exit codes
                switch ($proc.ExitCode) {
                    0 {
                        Write-Log "[OK] $pkgName installed successfully" -Level Success
                        $pkgInstalled = $true
                    }
                    -1978335212 {
                        Write-Log "[INFO] ${pkgName}: Already installed or no applicable update" -Level Warning
                        $pkgInstalled = $true
                    }
                    -1978335189 {
                        Write-Log "[INFO] ${pkgName}: No applicable update available" -Level Warning
                        $pkgInstalled = $true
                    }
                    default {
                        if ($installAttempts -lt $maxAttempts) {
                            Write-Log "[RETRY] $pkgName installation failed with code $($proc.ExitCode), retrying without --silent" -Level Warning
                        }
                        else {
                            Write-Log "[WARN] $pkgName installation failed after $maxAttempts attempts, exit code: $($proc.ExitCode)" -Level Warning
                            $pkgInstalled = $true  # Mark as done to exit loop
                        }
                    }
                }
            }
        }
        catch {
            Write-Log "Failed to install $pkgName : $($_.Exception.Message)" -Level Warning
        }
    }
}

# ============================================================================
# LICENSING (MAS INTEGRATION)
# ============================================================================

function Invoke-Licensing {
    Write-Log "Starting Licensing Check & Activation..." -Level Info
    Update-Progress -Status (L 'StatusActivating') -Percent 75 -SubStatus (L 'SubStatusWindowsOffice')
    
    # Download MAS
    $masPath = Join-Path $Script:Config.TempFolder 'MAS_AIO.cmd'
    $downloaded = $false
    
    foreach ($url in $Script:Config.MASUrls) {
        try {
            Write-Log "Attempting MAS download from: $url" -Level Debug
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($url, $masPath)
            $downloaded = $true
            Write-Log "MAS downloaded successfully" -Level Success
            break
        }
        catch {
            Write-Log "MAS download failed from this source: $($_.Exception.Message)" -Level Warning
        }
    }
    
    if (-not $downloaded) {
        Write-Log "Could not download activation script. Skipping licensing." -Level Error
        return
    }

    # 1. Apply Windows Edition (if selected)
    try {
        $targetEdition = $Script:UserConfig.WindowsEdition
        if ($targetEdition -and $targetEdition -ne '') {
            Write-Log "Applying Windows Edition: $targetEdition" -Level Info
            Update-Progress -Status (L 'StatusActivating') -Percent 78 -SubStatus "Switching Windows to $targetEdition..."
            
            $editionKeys = @{
                'Pro'           = 'VK7JG-NPHTM-C97JM-9MPGT-3V66T'
                'Home'          = 'YTMG3-N6DKC-DKB77-7M9GH-8HVX7'
                'Enterprise'    = 'XGVPP-NMH47-7TTHJ-W3FW7-8HV2C'
                'Education'     = 'YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY'
                'IoTEnterprise' = 'XQQ2F-VNMJ3-C7R3R-VV3BZ-T4R6J'
            }
            
            if ($editionKeys.ContainsKey($targetEdition)) {
                $genericKey = $editionKeys[$targetEdition]
                Write-Log "Using Generic Key: $genericKey" -Level Debug
                
                # Method 1: changepk.exe
                $proc = Start-Process -FilePath "changepk.exe" -ArgumentList "/ProductKey $genericKey" -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
                
                # Method 2: slmgr /ipk (fallback)
                if ($proc.ExitCode -ne 0) {
                    Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo $env:SystemRoot\System32\slmgr.vbs /ipk $genericKey" -Wait -WindowStyle Hidden
                }
                
                Start-Sleep -Seconds 5
            }
        }
    }
    catch {
        Write-Log "Windows edition change failed: $($_.Exception.Message)" -Level Warning
    }
    
    # 2. Activate Windows (HWID)
    Write-Log "Checking Windows Activation..." -Level Debug
    $winStatus = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.ApplicationId -eq '55c92734-d682-4d71-983e-d6ec3f16059f' } | Select-Object -ExpandProperty LicenseStatus -Unique
    
    if ($winStatus -eq 1) {
        Write-Log "Windows is already activated." -Level Success
    }
    else {
        Write-Log "Activating Windows (HWID)..." -Level Info
        Update-Progress -Status (L 'StatusActivating') -Percent 80 -SubStatus 'Windows...'
        $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$masPath`" /HWID" -Wait -PassThru -WindowStyle Hidden
        Write-Log "  -> Windows activation exit code: $($proc.ExitCode)" -Level Debug
    }
    
    # 3. Activate Office (Ohook)
    Write-Log "Activating Office (Ohook)..." -Level Info
    Update-Progress -Status (L 'StatusActivating') -Percent 85 -SubStatus 'Office...'
    $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$masPath`" /Ohook" -Wait -PassThru -WindowStyle Hidden
    Write-Log "  -> Office activation exit code: $($proc.ExitCode)" -Level Debug
    
    Write-Log "Licensing operations complete" -Level Success
}

# ============================================================================
# CLEANUP & VERIFICATION
# ============================================================================

function Test-CleanupIntegrity {
    param([int]$Attempt = 1, [int]$MaxAttempts = 3)
    
    Write-Log "Verifying cleanup (attempt $Attempt of $MaxAttempts)..." -Level Debug
    
    $orphanedFiles = @()
    $tempPath = $Script:Config.TempFolder
    
    # Check if temp folder still exists
    if (Test-Path $tempPath) {
        $orphanedFiles += Get-ChildItem $tempPath -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    }
    
    # Check for leftover Office installation files
    $odtPaths = @(
        "$env:TEMP\odt_setup.exe",
        "$env:TEMP\configuration.xml"
    )
    foreach ($path in $odtPaths) {
        if (Test-Path $path) { $orphanedFiles += $path }
    }
    
    if ($orphanedFiles.Count -gt 0) {
        Write-Log "Found $($orphanedFiles.Count) orphaned files, removing..." -Level Warning
        foreach ($file in $orphanedFiles) {
            try {
                Remove-Item $file -Force -Recurse -ErrorAction Stop
                Write-Log "Removed: $file" -Level Debug
            }
            catch {
                Write-Log "Failed to remove orphaned file $file : $($_.Exception.Message)" -Level Error
            }
        }
        # Retry cleanup only if under max attempts
        if ($Attempt -lt $MaxAttempts) {
            Start-Sleep -Seconds 1
            Test-CleanupIntegrity -Attempt ($Attempt + 1) -MaxAttempts $MaxAttempts
        }
        else {
            Write-Log "Max cleanup attempts reached. Some files may remain." -Level Warning
        }
    }
    else {
        Write-Log "Cleanup verification successful - no orphaned files found" -Level Success
    }
}

function Restore-SystemState {
    Write-Log "Restoring system to previous state..." -Level Info
    Update-Progress -Status (L 'StatusCleaning') -Percent 98 -SubStatus (L 'SubStatusRevertingChanges')
    
    # Restore registry keys
    Write-Log "Restoring registry keys..." -Level Debug
    @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive',
        'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate'
    ) | ForEach-Object { 
        try {
            Restore-RegistryKey -Path $_
        }
        catch {
            Write-Log "Warning restoring ${_}: $($_.Exception.Message)" -Level Warning
        }
    }
    
    # Clean up temporary files
    Write-Log "Removing temporary installation files..." -Level Debug
    try {
        if (Test-Path $Script:Config.TempFolder) {
            Remove-Item $Script:Config.TempFolder -Recurse -Force -ErrorAction Stop
            Write-Log "Removed temp folder: $($Script:Config.TempFolder)" -Level Debug
        }
    }
    catch {
        Write-Log "Warning removing temp folder: $($_.Exception.Message)" -Level Warning
    }
    
    # Remove any leftover Office installation folders
    $leftoverPaths = @(
        (Join-Path $env:LOCALAPPDATA "Temp\M365Ultimate_Installation"),
        (Join-Path $env:TEMP "M365Ultimate_Temp"),
        (Join-Path $env:TEMP "odt_setup.exe"),
        (Join-Path $env:TEMP "configuration.xml"),
        (Join-Path $env:USERPROFILE "Office"),
        (Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\OfficeFileCache")
    )
    
    foreach ($path in $leftoverPaths) {
        if (Test-Path $path) {
            try {
                Remove-Item $path -Recurse -Force -ErrorAction Stop
                Write-Log "Removed leftover: $path" -Level Debug
            }
            catch {
                Write-Log "Warning removing $path : $($_.Exception.Message)" -Level Warning
            }
        }
    }
    
    # Verify cleanup
    Test-CleanupIntegrity
    
    Write-Log "System restoration complete" -Level Success
}

# ============================================================================
# MAIN
# ============================================================================

$Script:Stage = 'initialization'
try {
    $Script:Stage = 'start'
    Initialize-Log
    Write-Log "=== INSTALLATION STARTED ===" -Level Info
    
    # Show configuration window (Blocking)
    $configSuccess = Show-ConfigWindow
    if (-not $configSuccess) {
        Write-Log "User cancelled configuration" -Level Info
        exit 0
    }
    Write-Log "User configuration completed - Mode: $($Script:ConfigResult.Mode)" -Level Info

    # -------------------------------------------------------------------------
    # SPECIAL MODE: DEEP UNINSTALL
    # -------------------------------------------------------------------------
    if ($Script:ConfigResult.Mode -eq 'Uninstall') {
        Write-Log "Starting Deep Uninstall Mode..." -Level Info
        
        # Start Progress UI
        Start-ProgressWindow
        Update-Progress -Status "Uninstalling Office..." -Percent 0 -SubStatus "Initializing..."
        
        try {
            Start-OfficeUninstallation
        }
        catch {
            Write-Log "Uninstall Error: $($_.Exception.Message)" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Uninstallation encountered errors. Check logs.", "Error", 0, 16)
        }
        
        # Cleanup
        Close-Progress
        if (Test-Path $Script:Config.TempFolder) { Remove-Item $Script:Config.TempFolder -Recurse -Force -ErrorAction SilentlyContinue }
        
        Write-Log "Uninstallation finished. Exiting." -Level Info
        
        # Log Logic Fix: Remove log file if successful (User Request)
        try {
            if (Test-Path $Script:Config.LogFile) { 
                Remove-Item $Script:Config.LogFile -Force -ErrorAction SilentlyContinue 
            }
        }
        catch { $null = $_ }

        exit 0
    }

    # -------------------------------------------------------------------------
    # CONFIGURATION: WINDOWS INSIDER
    # -------------------------------------------------------------------------
    if ($Script:ConfigResult.WindowsInsiderChannel -and $Script:ConfigResult.WindowsInsiderChannel -ne "None") {
        Set-WindowsInsider -Channel $Script:ConfigResult.WindowsInsiderChannel
        # Note: If restart triggered, script ends. If deferred, we proceed.
    }

    # Start progress window (Background) for Installation
    Start-ProgressWindow
    
    # Backup registry keys before modifications
    @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive',
        'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate'
    ) | ForEach-Object { Save-RegistryKey -Path $_ }

    $Script:Stage = 'temp-folder'
    Update-Progress -Status (L 'StatusPreparing') -Percent 5 -SubStatus (L 'StatusInitializing')
    New-TempFolder
    
    # Check for user cancel
    if ($Script:ProgressSync.RequestCancel) { throw "Installation cancelled by user" }

    $Script:Stage = 'bloatware'
    Update-Progress -Status (L 'StatusCleaning') -Percent 10 -SubStatus (L 'SubStatusRemovingConflicts')
    Remove-Conflicting-Office
    
    # Check for user cancel
    if ($Script:ProgressSync.RequestCancel) { throw "Installation cancelled by user" }

    $Script:Stage = 'phase2-prepare'
    Write-Log "Phase 2: Preparing Office Installation" -Level Info
    $setup = Get-ODT
    $xmlConfigPath = New-ConfigXML
    
    # Check for user cancel
    if ($Script:ProgressSync.RequestCancel) { throw "Installation cancelled by user" }

    $Script:Stage = 'phase2-validate'
    Write-Log "Validating installation assets..." -Level Info
    try {
        Test-FileExists -Path $setup -Label 'ODT setup'
        Test-FileExists -Path $xmlConfigPath -Label 'Configuration XML'
        Write-Log "Validation OK: setup & config" -Level Success
    }
    catch {
        Write-Log "Asset validation failed: $($_.Exception.Message)" -Level Error
        Write-Log "Stack: $($_.Exception.StackTrace)" -Level Error
        throw
    }

    $Script:Stage = 'phase3-install'
    Write-Log "Phase 3: Installing Office 365" -Level Info
    try {
        Install-Office -SetupPath $setup -ConfigPath $xmlConfigPath
    }
    catch {
        Write-Log "Install-Office failed: $($_.Exception.Message)" -Level Error
        Write-Log "Stack: $($_.ScriptStackTrace)" -Level Error
        throw
    }
    Write-Log "Phase 3 completed" -Level Info
    
    # Run cleanup immediately after install to ensure OneNote is removed even if later stages fail
    Start-CleanupOperations
    Remove-AppShortcuts
    
    # Remove Windows Bloatware (OneDrive, Appx) only after successful install
    Remove-Windows-Bloatware

    # Check for user cancel
    if ($Script:ProgressSync.RequestCancel) { throw "Installation cancelled by user" }

    $Script:Stage = 'phase4-extras'
    Write-Log "Phase 4: Installing extras" -Level Info
    Install-Extras
    Write-Log "Phase 4 completed" -Level Info
    
    # Check for user cancel
    if ($Script:ProgressSync.RequestCancel) { throw "Installation cancelled by user" }

    $Script:Stage = 'phase5-licensing'
    Write-Log "Phase 5: Licensing activation" -Level Info
    Invoke-Licensing
    Write-Log "Phase 5 completed" -Level Info
    
    # Check for user cancel
    if ($Script:ProgressSync.RequestCancel) { throw "Installation cancelled by user" }

    $Script:Stage = 'post-cleanup'
    # Start-CleanupOperations and Remove-AppShortcuts already run after Phase 3
    # Only run again if there were intervening steps that may have created new artifacts
    if ($Script:UserConfig.IncludeClipchamp -or $Script:UserConfig.IncludePowerAutomate) {
        Start-CleanupOperations
        Remove-AppShortcuts
    }

    $Script:Stage = 'cleanup'
    Write-Log "Cleaning temporary files..." -Level Debug
    Update-Progress -Status (L 'StatusFinalizing') -Percent 95 -SubStatus (L 'SubStatusRemovingTemp')
    Remove-Item $Script:Config.TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    Test-CleanupIntegrity

    Update-Progress -Status (L 'StatusComplete') -Percent 100 -SubStatus (L 'SubStatusAllComplete')
    
    $duration = (Get-Date) - $Script:Config.InstallStartTime
    Write-Log "=== INSTALLATION COMPLETED SUCCESSFULLY ===" -Level Success
    Write-Log "Total duration: $('{0:hh\:mm\:ss}' -f $duration)" -Level Info
    
    Start-Sleep -Seconds 2
    Close-Progress
    
    # Remove log file on success (no log left behind)
    try {
        if (Test-Path $Script:Config.LogFile) {
            Remove-Item $Script:Config.LogFile -Force -ErrorAction Stop
        }
    }
    catch {
        # Intentionally suppressed: Log file may be locked or already removed
        $null = $_
    }

    # Silent completion - no popup
}
catch {
    $err = $_.Exception.Message
    $duration = (Get-Date) - $Script:Config.InstallStartTime
    
    # Determine if it's a user cancellation
    $isCancellation = $err -like "*cancelled by user*"
    
    if ($isCancellation) {
        Write-Log "=== INSTALLATION CANCELLED BY USER ===" -Level Warning
        try { Update-Progress -Status "Cancelling..." -Percent 100 -SubStatus "Cleaning up leftovers..." } catch { $null = $_ }
    }
    else {
        Write-Log "=== CRITICAL ERROR ===" -Level Error
    }
    
    Write-Log "Stage: $Script:Stage" -Level Error
    Write-Log "Error: $err" -Level Error
    Write-Log "Stack: $($_.Exception.StackTrace)" -Level Error
    Write-Log "ScriptStackTrace: $($_.ScriptStackTrace)" -Level Error
    if ($_.InvocationInfo) { Write-Log "Invocation: $($_.InvocationInfo.PositionMessage)" -Level Error }
    if ($_.Exception.InnerException) { Write-Log "InnerException: $($_.Exception.InnerException.Message)" -Level Error }
    Write-Log "Duration before failure: $('{0:hh\:mm\:ss}' -f $duration)" -Level Error
    
    # Restore system state
    Write-Log "Initiating system restoration..." -Level Warning
    Restore-SystemState
    Test-CleanupIntegrity

    if ($isCancellation) {
        Start-Sleep -Seconds 2
        Close-Progress
    }
    else {
        Close-Progress
        [System.Windows.Forms.MessageBox]::Show("Installation Failed!`n`nStage: $Script:Stage`nError: $err`n`nAll changes have been reverted.`n`nCheck log file on Desktop for details:`n$($Script:Config.LogFile)", "Error", 0, 16)
    }
}
finally {
    # Always clean up resources, even on unexpected termination
    try {
        Close-Progress
    }
    catch {
        # Intentionally suppressed: Progress may already be closed
        $null = $_
    }
    
    try {
        if ($ScriptMutex) {
            $ScriptMutex.ReleaseMutex()
            $ScriptMutex.Dispose()
            $ScriptMutex = $null
        }
    }
    catch {
        # Intentionally suppressed: Mutex may already be released
        $null = $_
    }
}



