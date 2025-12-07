# Microsoft 365 Ultimate Installer

## Overview
The **Microsoft 365 Ultimate Installer** is a robust, automated PowerShell solution designed for AC Tech to streamline the deployment of Microsoft 365 Enterprise and related applications. It features a modern WPF-based Graphical User Interface (GUI) for configuration and progress tracking.

## Features
- **User-Friendly GUI**: 
  - **Configuration Window**: Allows selection of Office version, languages, and specific applications.
  - **Progress Window**: Non-blocking background window with real-time status updates, minimize/maximize controls, and cancellation support.
- **Customizable Installation**:
  - **Express Mode**: Quickly installs the standard suite (Word, Excel, PowerPoint, Outlook, Teams, Clipchamp).
  - **Custom Mode**: Granular control over every application (Access, Publisher, Project, Visio, etc.).
- **Multi-Language Support**: Supports 29+ languages including English, Portuguese, Spanish, German, French, Japanese, and more.
- **Automated Cleanup**: Automatically removes conflicting Office installations and cleans up temporary files after installation.
- **Resilience**: Includes robust error handling, logging, and a rollback mechanism in case of cancellation or failure.

## Prerequisites
- Windows 10 or Windows 11 (64-bit recommended).
- PowerShell 5.1 or later.
- Internet connection (for downloading Office Deployment Tool and installation files).
- Administrator privileges.

## Usage
1. Right-click `Microsoft 365 Ultimate Installer.ps1`.
2. Select **Run with PowerShell**.
3. Follow the on-screen prompts to configure your installation.

## Directory Structure
The project is self-contained, but during execution, it utilizes:
- **Desktop**: Log file location.
- **%LOCALAPPDATA%\Temp\M365Ultimate_Installation**: Temporary working directory for ODT and downloads.

## Technical Details
- **Language**: PowerShell
- **UI Framework**: WPF (Windows Presentation Foundation) via XAML.
- **Concurrency**: Uses PowerShell Runspaces to keep the UI responsive during long-running installation processes.
- **Deployment Method**: Uses the official Microsoft Office Deployment Tool (ODT).

## License
Private Property of AC Tech.
