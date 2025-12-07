# Microsoft 365 Ultimate Installer - Project Rules & Guidelines

This document outlines the coding standards, architectural decisions, and best practices for the **Microsoft 365 Ultimate Installer** project.

## 1. Project Architecture
- **Single-File Distribution**: The script must remain a single `.ps1` file for ease of distribution. All helper functions, XAML definitions, and logic must be contained within.
- **UI Separation**: 
  - **Configuration Window**: Runs in the main thread (blocking) to gather user input.
  - **Progress Window**: MUST run in a separate **Runspace** (background thread) to ensure the UI remains responsive while the installation (which is blocking) occurs.
- **WPF/XAML**: The UI is built using Windows Presentation Foundation (WPF) loaded via `[Windows.Markup.XamlReader]`.

## 2. Coding Standards (PowerShell)
- **Error Handling**: 
  - Use `try/catch` blocks for all critical operations (file I/O, network requests, process execution).
  - Set `$ErrorActionPreference = 'Stop'` at the beginning of the script.
- **Logging**: 
  - **DO NOT** use `Write-Host` for critical information. Use the custom `Write-Log` function which writes to both the console and a log file.
  - Log file location: User's Desktop (`Microsoft 365 Ultimate Installer.log`).
- **Variables**:
  - Avoid using automatic variables like `$sender` or `$args` in event handlers if possible, or rename them to avoid PSScriptAnalyzer warnings (e.g., use `$s` instead of `$sender`).
  - Use `$Script:` scope for variables shared between functions or runspaces.

## 3. UI Guidelines
- **XAML Namespaces**: Always include `xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"` in the `<Window>` tag to support `x:Name`.
- **Responsiveness**: The Progress Window must support Minimize and Maximize actions.
- **Cancellation**: 
  - The "Cancel" button should not immediately close the window. It should set a flag (`$Sync.RequestCancel`) and trigger a cleanup/rollback process.
  - The window should only close after cleanup is complete.

## 4. Installation Logic
- **Office Deployment Tool (ODT)**: The script downloads the latest ODT at runtime. Do not bundle binaries.
- **Configuration XML**: Generated dynamically based on user selection.
- **Clean Install**: Always check for and remove conflicting Office installations (Click-to-Run) before starting.

## 5. Version Control
- **Private Repository**: This project is hosted in a private GitHub repository under AC Tech.
- **Sensitive Data**: Do not commit secrets, API keys, or internal network paths.

## 6. VS Code Configuration
- Ensure the `PowerShell` extension is enabled.
- Use `PSScriptAnalyzer` for linting.
