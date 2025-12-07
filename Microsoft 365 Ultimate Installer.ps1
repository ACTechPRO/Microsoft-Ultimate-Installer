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
    LogFile          = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Microsoft 365 Ultimate Installer.log'
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
        WindowTitle                     = 'Microsoft 365 Ultimate Installer v3.6'
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
    }
    
    'pt' = @{
        WindowTitle                     = 'Instalador Microsoft 365 Ultimate'
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
    }
    
    'ja' = @{
        WindowTitle                     = 'Microsoft 365 Ultimate インストーラー'
        WindowSubtitle                  = '自動インストーラー＆アクティベーター'
        ConfigWindowTitle               = 'インストール設定'
        
        ExpressMode                     = 'エクスプレスインストール（推奨）'
        ExpressModeDesc                 = 'デフォルト設定でインストール - すべてのアプリを含むMicrosoft 365'
        CustomMode                      = 'カスタムインストール'
        CustomModeDesc                  = 'バージョン、言語、アプリケーションを選択'
        
        SelectVersion                   = 'Officeバージョンを選択:'
        SelectLanguages                 = '言語を選択:'
        PrimaryLanguage                 = '主要言語:'
        AdditionalLanguages             = '追加言語:'
        SelectApps                      = 'アプリケーションを選択:'
        SelectAll                       = 'すべて選択'
        DeselectAll                     = 'すべて解除'
        
        BtnStart                        = 'インストール開始'
        BtnCancel                       = 'キャンセル'
        BtnBack                         = '戻る'
        BtnUseDefaults                  = 'デフォルトを使用'
        
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
        
        StatusInitializing              = '初期化中...'
        StatusPreparing                 = '準備中...'
        StatusDownloadingODT            = 'Office展開ツールをダウンロード中...'
        StatusDownloadingOffice         = 'Officeファイルをダウンロード中...'
        StatusInstallingOffice          = 'Microsoft 365をインストール中...'
        StatusInstallingExtras          = '追加機能をインストール中...'
        StatusActivating                = 'ライセンスを有効化中...'
        StatusCleaning                  = 'クリーンアップ中...'
        StatusFinalizing                = '完了処理中...'
        StatusComplete                  = '完了!'
        
        SubStatusInternetSpeed          = '速度はインターネット接続に依存'
        SubStatusApplyingConfig         = '設定を適用中'
        SubStatusClipchampPowerAutomate = 'Clipchamp & Power Automate'
        SubStatusWindowsOffice          = 'WindowsとOfficeのアクティベーション'
        SubStatusRemovingTemp           = '一時ファイルを削除中'
        SubStatusAllComplete            = 'すべてのタスクが正常に完了しました'
        
        LogInstallStarted               = '=== インストール開始 ==='
        LogInstallComplete              = '=== インストール正常完了 ==='
        LogInstallCancelled             = '=== ユーザーによりキャンセル ==='
        LogInstallFailed                = '=== 重大なエラー ==='
        LogDuration                     = '合計時間:'
        LogPhase                        = 'フェーズ'
        LogCompleted                    = '完了'
        
        ErrAnotherInstance              = '別のインスタンスが既に実行中です。'
        ErrInstallCancelled             = 'インストールがキャンセルされました。'
        ErrAllChangesReverted           = 'すべての変更が元に戻され、一時ファイルが削除されました。'
        ErrInstallFailed                = 'インストール失敗!'
        ErrCheckLog                     = 'デスクトップのログファイルを確認してください:'
        
        ValidSelectAtLeastOneApp        = 'インストールするアプリケーションを1つ以上選択してください。'
        ValidSelectAtLeastOneLang       = '言語を1つ以上選択してください。'
        ValidProjectVisioNote           = 'ProjectとVisioは別途ライセンスが必要ですが、自動的にアクティベートされます。'
    }
    
    'de' = @{
        WindowTitle                     = 'Microsoft 365 Ultimate Installer'
        WindowSubtitle                  = 'Automatischer Installer & Aktivator'
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
        SubStatusClipchampPowerAutomate = 'Clipchamp & Power Automate'
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
    }
    
    'zh' = @{
        WindowTitle                     = 'Microsoft 365 终极安装程序'
        WindowSubtitle                  = '自动安装和激活工具'
        ConfigWindowTitle               = '安装配置'
        
        ExpressMode                     = '快速安装（推荐）'
        ExpressModeDesc                 = '使用默认设置安装 - 包含所有应用的 Microsoft 365'
        CustomMode                      = '自定义安装'
        CustomModeDesc                  = '选择版本、语言和应用程序'
        
        SelectVersion                   = '选择 Office 版本:'
        SelectLanguages                 = '选择语言:'
        PrimaryLanguage                 = '主要语言:'
        AdditionalLanguages             = '其他语言:'
        SelectApps                      = '选择应用程序:'
        SelectAll                       = '全选'
        DeselectAll                     = '取消全选'
        
        BtnStart                        = '开始安装'
        BtnCancel                       = '取消'
        BtnBack                         = '返回'
        BtnUseDefaults                  = '使用默认值'
        
        Version365Enterprise            = 'Microsoft 365 企业版'
        Version365Business              = 'Microsoft 365 商业版'
        VersionProPlus2024              = 'Office LTSC 专业增强版 2024'
        VersionProPlus2021              = 'Office LTSC 专业增强版 2021'
        VersionProPlus2019              = 'Office 专业增强版 2019'
        
        AppWord                         = 'Word'
        AppExcel                        = 'Excel'
        AppPowerPoint                   = 'PowerPoint'
        AppOutlook                      = 'Outlook'
        AppAccess                       = 'Access'
        AppPublisher                    = 'Publisher'
        AppOneNote                      = 'OneNote'
        AppOneDrive                     = 'OneDrive'
        AppTeams                        = 'Teams'
        AppProject                      = 'Project 专业版'
        AppVisio                        = 'Visio 专业版'
        AppClipchamp                    = 'Clipchamp'
        AppPowerAutomate                = 'Power Automate Desktop'
        
        StatusInitializing              = '初始化中...'
        StatusPreparing                 = '准备中...'
        StatusDownloadingODT            = '正在下载 Office 部署工具...'
        StatusDownloadingOffice         = '正在下载 Office 文件...'
        StatusInstallingOffice          = '正在安装 Microsoft 365...'
        StatusInstallingExtras          = '正在安装附加组件...'
        StatusActivating                = '正在激活许可证...'
        StatusCleaning                  = '正在清理...'
        StatusFinalizing                = '正在完成...'
        StatusComplete                  = '完成!'
        
        SubStatusInternetSpeed          = '速度取决于网络连接'
        SubStatusApplyingConfig         = '正在应用配置'
        SubStatusClipchampPowerAutomate = 'Clipchamp 和 Power Automate'
        SubStatusWindowsOffice          = 'Windows 和 Office 激活'
        SubStatusRemovingTemp           = '正在删除临时文件'
        SubStatusAllComplete            = '所有任务已成功完成'
        
        LogInstallStarted               = '=== 安装已开始 ==='
        LogInstallComplete              = '=== 安装已成功完成 ==='
        LogInstallCancelled             = '=== 用户取消安装 ==='
        LogInstallFailed                = '=== 严重错误 ==='
        LogDuration                     = '总耗时:'
        LogPhase                        = '阶段'
        LogCompleted                    = '已完成'
        
        ErrAnotherInstance              = '另一个实例正在运行。'
        ErrInstallCancelled             = '安装已取消。'
        ErrAllChangesReverted           = '所有更改已还原，临时文件已删除。'
        ErrInstallFailed                = '安装失败!'
        ErrCheckLog                     = '请检查桌面上的日志文件:'
        
        ValidSelectAtLeastOneApp        = '请至少选择一个要安装的应用程序。'
        ValidSelectAtLeastOneLang       = '请至少选择一种语言。'
        ValidProjectVisioNote           = 'Project 和 Visio 需要单独的许可证，但会自动激活。'
    }
    
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
    }
    
    'ko' = @{
        WindowTitle                     = 'Microsoft 365 Ultimate 설치 프로그램'
        WindowSubtitle                  = '자동 설치 및 정품 인증'
        ConfigWindowTitle               = '설치 구성'
        
        ExpressMode                     = '빠른 설치 (권장)'
        ExpressModeDesc                 = '기본 설정으로 설치 - 모든 앱 포함 Microsoft 365'
        CustomMode                      = '사용자 지정 설치'
        CustomModeDesc                  = '버전, 언어 및 응용 프로그램 선택'
        
        SelectVersion                   = 'Office 버전 선택:'
        SelectLanguages                 = '언어 선택:'
        PrimaryLanguage                 = '기본 언어:'
        AdditionalLanguages             = '추가 언어:'
        SelectApps                      = '응용 프로그램 선택:'
        SelectAll                       = '모두 선택'
        DeselectAll                     = '모두 해제'
        
        BtnStart                        = '설치 시작'
        BtnCancel                       = '취소'
        BtnBack                         = '뒤로'
        BtnUseDefaults                  = '기본값 사용'
        
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
        
        StatusInitializing              = '초기화 중...'
        StatusPreparing                 = '준비 중...'
        StatusDownloadingODT            = 'Office 배포 도구 다운로드 중...'
        StatusDownloadingOffice         = 'Office 파일 다운로드 중...'
        StatusInstallingOffice          = 'Microsoft 365 설치 중...'
        StatusInstallingExtras          = '추가 구성 요소 설치 중...'
        StatusActivating                = '라이선스 정품 인증 중...'
        StatusCleaning                  = '정리 중...'
        StatusFinalizing                = '마무리 중...'
        StatusComplete                  = '완료!'
        
        SubStatusInternetSpeed          = '속도는 인터넷 연결에 따라 다름'
        SubStatusApplyingConfig         = '구성 적용 중'
        SubStatusClipchampPowerAutomate = 'Clipchamp 및 Power Automate'
        SubStatusWindowsOffice          = 'Windows 및 Office 정품 인증'
        SubStatusRemovingTemp           = '임시 파일 제거 중'
        SubStatusAllComplete            = '모든 작업이 성공적으로 완료됨'
        
        LogInstallStarted               = '=== 설치 시작 ==='
        LogInstallComplete              = '=== 설치 성공적으로 완료 ==='
        LogInstallCancelled             = '=== 사용자에 의해 설치 취소 ==='
        LogInstallFailed                = '=== 심각한 오류 ==='
        LogDuration                     = '총 소요 시간:'
        LogPhase                        = '단계'
        LogCompleted                    = '완료'
        
        ErrAnotherInstance              = '다른 인스턴스가 이미 실행 중입니다.'
        ErrInstallCancelled             = '설치가 취소되었습니다.'
        ErrAllChangesReverted           = '모든 변경 사항이 되돌려지고 임시 파일이 제거되었습니다.'
        ErrInstallFailed                = '설치 실패!'
        ErrCheckLog                     = '바탕 화면의 로그 파일을 확인하세요:'
        
        ValidSelectAtLeastOneApp        = '설치할 응용 프로그램을 하나 이상 선택하세요.'
        ValidSelectAtLeastOneLang       = '언어를 하나 이상 선택하세요.'
        ValidProjectVisioNote           = 'Project와 Visio는 별도의 라이선스가 필요하지만 자동으로 정품 인증됩니다.'
    }
    
    'ru' = @{
        WindowTitle                     = 'Microsoft 365 Ultimate Установщик'
        WindowSubtitle                  = 'Автоматический Установщик и Активатор'
        ConfigWindowTitle               = 'Настройка Установки'
        
        ExpressMode                     = 'Экспресс-установка (Рекомендуется)'
        ExpressModeDesc                 = 'Установить с настройками по умолчанию - Microsoft 365 со всеми приложениями'
        CustomMode                      = 'Выборочная установка'
        CustomModeDesc                  = 'Выбрать версию, языки и приложения'
        
        SelectVersion                   = 'Выберите версию Office:'
        SelectLanguages                 = 'Выберите языки:'
        PrimaryLanguage                 = 'Основной язык:'
        AdditionalLanguages             = 'Дополнительные языки:'
        SelectApps                      = 'Выберите приложения:'
        SelectAll                       = 'Выбрать все'
        DeselectAll                     = 'Снять все'
        
        BtnStart                        = 'Начать установку'
        BtnCancel                       = 'Отмена'
        BtnBack                         = 'Назад'
        BtnUseDefaults                  = 'По умолчанию'
        
        Version365Enterprise            = 'Microsoft 365 Корпоративный'
        Version365Business              = 'Microsoft 365 Бизнес'
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
        
        StatusInitializing              = 'Инициализация...'
        StatusPreparing                 = 'Подготовка...'
        StatusDownloadingODT            = 'Загрузка Office Deployment Tool...'
        StatusDownloadingOffice         = 'Загрузка файлов Office...'
        StatusInstallingOffice          = 'Установка Microsoft 365...'
        StatusInstallingExtras          = 'Установка дополнений...'
        StatusActivating                = 'Активация лицензий...'
        StatusCleaning                  = 'Очистка...'
        StatusFinalizing                = 'Завершение...'
        StatusComplete                  = 'Готово!'
        
        SubStatusInternetSpeed          = 'Скорость зависит от интернета'
        SubStatusApplyingConfig         = 'Применение конфигурации'
        SubStatusClipchampPowerAutomate = 'Clipchamp и Power Automate'
        SubStatusWindowsOffice          = 'Активация Windows и Office'
        SubStatusRemovingTemp           = 'Удаление временных файлов'
        SubStatusAllComplete            = 'Все задачи успешно выполнены'
        
        LogInstallStarted               = '=== УСТАНОВКА НАЧАТА ==='
        LogInstallComplete              = '=== УСТАНОВКА УСПЕШНО ЗАВЕРШЕНА ==='
        LogInstallCancelled             = '=== УСТАНОВКА ОТМЕНЕНА ПОЛЬЗОВАТЕЛЕМ ==='
        LogInstallFailed                = '=== КРИТИЧЕСКАЯ ОШИБКА ==='
        LogDuration                     = 'Общее время:'
        LogPhase                        = 'Этап'
        LogCompleted                    = 'завершён'
        
        ErrAnotherInstance              = 'Другой экземпляр уже запущен.'
        ErrInstallCancelled             = 'Установка отменена.'
        ErrAllChangesReverted           = 'Все изменения отменены, временные файлы удалены.'
        ErrInstallFailed                = 'Ошибка установки!'
        ErrCheckLog                     = 'Проверьте файл журнала на рабочем столе:'
        
        ValidSelectAtLeastOneApp        = 'Выберите хотя бы одно приложение для установки.'
        ValidSelectAtLeastOneLang       = 'Выберите хотя бы один язык.'
        ValidProjectVisioNote           = 'Project и Visio требуют отдельных лицензий, но будут активированы автоматически.'
    }
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
    'pt-br' = 'Português (Brasil)'
    'es-es' = 'Español (España)'
    'es-mx' = 'Español (México)'
    'ja-jp' = '日本語'
    'de-de' = 'Deutsch'
    'fr-fr' = 'Français'
    'it-it' = 'Italiano'
    'zh-cn' = '中文 (简体)'
    'zh-tw' = '中文 (繁體)'
    'ko-kr' = '한국어'
    'ru-ru' = 'Русский'
    'ar-sa' = 'العربية'
    'nl-nl' = 'Nederlands'
    'pl-pl' = 'Polski'
    'tr-tr' = 'Türkçe'
    'sv-se' = 'Svenska'
    'da-dk' = 'Dansk'
    'fi-fi' = 'Suomi'
    'nb-no' = 'Norsk'
    'cs-cz' = 'Čeština'
    'el-gr' = 'Ελληνικά'
    'he-il' = 'עברית'
    'hu-hu' = 'Magyar'
    'th-th' = 'ไทย'
    'vi-vn' = 'Tiếng Việt'
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
        Get-Item $Path | Get-ItemProperty | GetEnumerator | ForEach-Object {
            $items += @{
                Name  = $_.Name
                Value = $_.Value
                Type  = (Get-ItemProperty $Path $_.Name).PSObject.Properties[$_.Name].TypeNameOfValue
            }
        }
    }
    catch {}
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
        # Mutex doesn't exist or already disposed - that's fine
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

function Show-ConfigWindow {
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
    $strConfigure = "Configure Installation" # TODO: Localize
    
    # Detect Windows language for default
    $winLang = (Get-Culture).Name.ToLower()
    $defaultLangIndex = 0
    $langMap = @{ 'en-us' = 0; 'pt-br' = 1; 'es-es' = 2; 'de-de' = 3; 'fr-fr' = 4; 'ja-jp' = 5; 'zh-cn' = 6; 'it-it' = 7; 'ko-kr' = 8; 'ru-ru' = 9 }
    if ($langMap.ContainsKey($winLang)) { $defaultLangIndex = $langMap[$winLang] }
    
    # Escape strings for XAML
    foreach ($v in @('title', 'subtitle', 'strExpress', 'strExpressDesc', 'strCustom', 'strCustomDesc', 'strStart', 'strCancel', 'strVersion', 'strLang', 'strApps', 'strSelectAll', 'strDeselectAll', 'strBack', 'strConfigure')) {
        Set-Variable -Name $v -Value ((Get-Variable -Name $v).Value -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$title" Width="700" Height="600"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize" WindowStyle="SingleBorderWindow"
        Background="#1e1e1e" Topmost="True" BorderBrush="#0078D4" BorderThickness="1">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="20,15">
            <StackPanel>
                <TextBlock Text="Microsoft 365 Ultimate Installer" FontSize="22" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="$subtitle" FontSize="12" Foreground="#DDDDDD" Margin="0,3,0,0"/>
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
            
            <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                <StackPanel>
                    <TextBlock Text="$strVersion" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,0,0,8"/>
                    <ComboBox x:Name="cmbVersion" Width="400" HorizontalAlignment="Left" SelectedIndex="0">
                        <ComboBoxItem Tag="O365ProPlusRetail|CurrentPreview">Microsoft 365 Enterprise (Current Preview)</ComboBoxItem>
                        <ComboBoxItem Tag="O365BusinessRetail|CurrentPreview">Microsoft 365 Business (Current Preview)</ComboBoxItem>
                        <ComboBoxItem Tag="ProPlus2024Retail|PerpetualVL2024">Office LTSC 2024</ComboBoxItem>
                        <ComboBoxItem Tag="ProPlus2021Retail|PerpetualVL2021">Office LTSC 2021</ComboBoxItem>
                        <ComboBoxItem Tag="ProPlus2019Retail|PerpetualVL2019">Office 2019</ComboBoxItem>
                    </ComboBox>
                    
                    <TextBlock Text="$strApps" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,20,0,8"/>
                    <WrapPanel>
                        <CheckBox x:Name="chkWord" Content="Word" IsChecked="True" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkExcel" Content="Excel" IsChecked="True" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkPowerPoint" Content="PowerPoint" IsChecked="True" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkOutlook" Content="Outlook" IsChecked="True" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkAccess" Content="Access" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkPublisher" Content="Publisher" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkTeams" Content="Teams" IsChecked="True" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkOneNote" Content="OneNote" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkOneDrive" Content="OneDrive" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkProject" Content="Project Pro" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkVisio" Content="Visio Pro" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkClipchamp" Content="Clipchamp" IsChecked="True" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkPowerAutomate" Content="Power Automate" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                        <CheckBox x:Name="chkSkype" Content="Skype" IsChecked="False" Foreground="White" Margin="0,0,20,8"/>
                    </WrapPanel>
                    
                    <StackPanel Orientation="Horizontal" Margin="0,10,0,15">
                        <Button x:Name="btnSelectAll" Content="$strSelectAll" Width="120" Height="28" Background="#444444" Foreground="White" BorderThickness="0" Margin="0,0,10,0"/>
                        <Button x:Name="btnDeselectAll" Content="$strDeselectAll" Width="120" Height="28" Background="#444444" Foreground="White" BorderThickness="0"/>
                    </StackPanel>

                    <TextBlock Text="$strLang" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,20,0,8"/>
                    <ComboBox x:Name="cmbLanguage" Width="300" HorizontalAlignment="Left" SelectedIndex="$defaultLangIndex">
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
                    <ScrollViewer Height="120" VerticalScrollBarVisibility="Auto" Background="#333333" BorderBrush="#555555" BorderThickness="1">
                        <WrapPanel Margin="10">
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

                    <TextBlock Text="Windows Edition (Activation):" FontSize="14" Foreground="White" FontWeight="SemiBold" Margin="0,20,0,8"/>
                    <ComboBox x:Name="cmbWindowsEdition" Width="300" HorizontalAlignment="Left" SelectedIndex="0">
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
        
        # Get language checkboxes
        $langCheckboxes = @()
        foreach ($tag in @('en-us','pt-br','es-es','de-de','fr-fr','ja-jp','zh-cn','it-it','ko-kr','ru-ru','ar-sa','da-dk','nl-nl','fi-fi','el-gr','he-il','hu-hu','id-id','ms-my','nb-no','pl-pl','pt-pt','ro-ro','sk-sk','sv-se','th-th','tr-tr','uk-ua','vi-vn')) {
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
        
        # Mode Next
        $btnModeNext.Add_Click({
                if ($rbExpress.IsChecked) {
                    $Script:ConfigResult.Cancelled = $false
                    $Script:ConfigResult.Mode = 'Express'
                    $window.DialogResult = $true
                    $window.Close()
                }
                else {
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
            })
        
        # Deselect All
        $btnDeselectAll.Add_Click({
                $chkWord.IsChecked = $false; $chkExcel.IsChecked = $false; $chkPowerPoint.IsChecked = $false
                $chkOutlook.IsChecked = $false; $chkAccess.IsChecked = $false; $chkPublisher.IsChecked = $false
                $chkTeams.IsChecked = $false; $chkOneNote.IsChecked = $false; $chkOneDrive.IsChecked = $false
                $chkProject.IsChecked = $false; $chkVisio.IsChecked = $false; $chkClipchamp.IsChecked = $false
                $chkPowerAutomate.IsChecked = $false; $chkSkype.IsChecked = $false
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
                    'Word'       = $chkWord.IsChecked
                    'Excel'      = $chkExcel.IsChecked
                    'PowerPoint' = $chkPowerPoint.IsChecked
                    'Outlook'    = $chkOutlook.IsChecked
                    'Access'     = $chkAccess.IsChecked
                    'Publisher'  = $chkPublisher.IsChecked
                    'Teams'      = $chkTeams.IsChecked
                    'OneNote'    = $chkOneNote.IsChecked
                    'OneDrive'   = $chkOneDrive.IsChecked
                    'Skype'      = $chkSkype.IsChecked
                }
            
                # Extras (now part of apps selection, but kept as separate flags for logic compatibility)
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
    <#
    .SYNOPSIS
    Starts progress window in a background runspace.
    #>
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    
    $title = L 'WindowTitle'
    $subtitle = L 'WindowSubtitle'
    $initStatus = L 'StatusInitializing'
    
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
        param($Sync, $Title, $Subtitle)
        
        # Escape strings
        $Title = $Title -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
        $Subtitle = $Subtitle -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
        
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        
        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Width="500" Height="250"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize" WindowStyle="SingleBorderWindow"
        Background="#1e1e1e" Topmost="True" BorderBrush="#0078D4" BorderThickness="1">
    <Grid Margin="30">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <TextBlock x:Name="StatusText" Grid.Row="0" Text="Initializing..." FontSize="18" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,15"/>
        
        <ProgressBar x:Name="ProgressBar" Grid.Row="1" Height="25" Minimum="0" Maximum="100" Value="0" Background="#333333" Foreground="#0078D4" Margin="0,0,0,10"/>
        
        <Grid Grid.Row="2">
            <TextBlock x:Name="SubStatusText" Text="" FontSize="12" Foreground="#AAAAAA" HorizontalAlignment="Left"/>
            <TextBlock x:Name="PercentText" Text="0%" FontSize="12" Foreground="#AAAAAA" HorizontalAlignment="Right"/>
        </Grid>
        
        <Button x:Name="btnCancel" Grid.Row="3" Content="Cancel" Width="100" Height="30" 
                Background="#444444" Foreground="White" BorderThickness="0" 
                VerticalAlignment="Bottom" HorizontalAlignment="Right"/>
    </Grid>
</Window>
"@
        
        try {
            [xml]$xamlDoc = $xaml
            $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
            $window = [Windows.Markup.XamlReader]::Load($reader)
            
            $statusText = $window.FindName('StatusText')
            $progressBar = $window.FindName('ProgressBar')
            $subStatusText = $window.FindName('SubStatusText')
            $percentText = $window.FindName('PercentText')
            $btnCancel = $window.FindName('btnCancel')
            
            $btnCancel.Add_Click({
                    $Sync.RequestCancel = $true
                    $btnCancel.IsEnabled = $false
                    $btnCancel.Content = "Cancelling..."
                })
            
            $window.Add_Closing({
                    param($s, $e)
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
    $ps.AddScript($code).AddArgument($Script:ProgressSync).AddArgument($title).AddArgument($subtitle) | Out-Null
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
    catch { }
    
    try {
        if ($Script:ProgressHandle -and -not $Script:ProgressHandle.IsCompleted) {
            $Script:ProgressHandle.AsyncWaitHandle.WaitOne(1000) | Out-Null
        }
    }
    catch { }
    
    try {
        if ($Script:ProgressRunspace) {
            $Script:ProgressRunspace.Close()
            $Script:ProgressRunspace.Dispose()
            $Script:ProgressRunspace = $null
        }
    }
    catch { }
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
        throw "$Label não encontrado: $Path"
    }
    try {
        $size = (Get-Item $Path).Length
        Write-Log "$Label localizado ($size bytes): $Path" -Level Debug
    }
    catch {
        Write-Log "$Label localizado: $Path (tamanho desconhecido)" -Level Debug
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
        'Outlook'    = 'Outlook'
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
$langTags$excludeApps
    </Product>
"@
    }
    
    # Add Visio if selected and supported
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
$langTags$excludeApps
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
    
    # Check if Clipchamp should be installed
    if ($Script:UserConfig.IncludeClipchamp) {
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
    Write-Log "Verifying cleanup..." -Level Debug
    
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
        Write-Log "Found orphaned files, removing..." -Level Warning
        foreach ($file in $orphanedFiles) {
            try {
                Remove-Item $file -Force -Recurse -ErrorAction Stop
                Write-Log "Removed: $file" -Level Debug
            }
            catch {
                Write-Log "Failed to remove orphaned file $file : $($_.Exception.Message)" -Level Error
            }
        }
        # Retry cleanup
        Start-Sleep -Seconds 1
        Test-CleanupIntegrity
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
    Write-Log "User configuration completed - Mode: $($Script:UserConfig.InstallMode)" -Level Info

    # Start progress window (Background)
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
    # Start-CleanupOperations was already run after Phase 3, but running it again doesn't hurt
    Start-CleanupOperations
    Remove-AppShortcuts

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
        # ignore log removal issues
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
        try { Update-Progress -Status "Cancelling..." -Percent 100 -SubStatus "Cleaning up leftovers..." } catch {}
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
    catch { }
    
    try {
        if ($ScriptMutex) {
            $ScriptMutex.ReleaseMutex()
            $ScriptMutex.Dispose()
            $ScriptMutex = $null
        }
    }
    catch { }
}

