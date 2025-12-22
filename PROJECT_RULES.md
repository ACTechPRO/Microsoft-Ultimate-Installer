# Microsoft Ultimate Installer - Project-Specific Rules

## Project Overview
- **Name**: Microsoft 365 Ultimate Installer
- **Repository**: https://github.com/ACTechPRO/Microsoft-Ultimate-Installer
- **Local Path**: `D:\Microsoft Ultimate Installer`
- **Type**: Monolithic PowerShell Script (`.ps1`).

## Core Principles
1.  **Single File Directive**: The entire logic MUST reside in one `.ps1` file. No external module dependencies (unless auto-downloaded).
2.  **Self-Elevation**: Script must auto-request Admin privileges.
3.  **Clean Execution**: Remove all temporary files (`C:\MSInstallerTemp`, etc.) upon checking exit.
4.  **UI/UX**: WPF-based GUI. Modern, clean, professional.
5.  **Language**: PT-BR default.

## Features
- **Office 365**: Installation, customization, activation (Massgrave method).
- **Tools**: Install VS Code, 7-Zip, AnyDesk, etc.
- **System**: Debloat Windows, Update Drivers.
- **Shortcuts**: Create reliable shortcuts for all installed apps.

## Tech Stack
- **Language**: PowerShell 5.1 / 7+ compatible.
- **GUI**: WPF (XAML embedded in PS1).
- **External**: Uses `irm ... | iex` for activation scripts (MAS).

## Assets
- **Header**: `Microsoft Ultimate Installer.png`
- **Footer**: `AC Tech Transparente Invertido.ico`
