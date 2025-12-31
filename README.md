<div align="center">

<img src="assets/Microsoft Ultimate Installer Black Background.png" alt="Microsoft Ultimate Installer Hero" width="100%" />

# Microsoft Ultimate Installer
### The Last Windows Setup Script You'll Ever Need.

![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![Version](https://img.shields.io/badge/Version-1.0.0-green?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-orange?style=for-the-badge)

</div>

---

## ⚡ Initiate Sequence (Admin)

Copy and paste this into a **PowerShell (Admin)** terminal. Watch the magic happen.

```powershell
irm ult.ac-tech.pro | iex
```

> **Note**: Requires Internet Connection. If blocked by execution policy, run `Set-ExecutionPolicy Unrestricted -Scope Process` first.

---

## 📡 Transmission: Philosophy & Overview

**Microsoft Ultimate Installer (MUI)** is a sophisticated automation framework designed for power users, developers, and sysadmins. Windows 11 is excellent, but it comes cluttered with "Consumer Experience" apps, telemetry, and redundant services that slow down your workflow.

MUI serves one purpose: **To give you a pristine, "Pro-Grade" environment in under 10 minutes.** It bypasses the bloat, automates the tedious, and handles licensing so you can focus on building.

### Why MUI?
*   **Zero-Click Visual Studio**: Installs Enterprise/Pro editions silently with your curated workloads. No more clicking "Next -> Next -> I Agree".
*   **Office 365 Automation**: Deploys the full suite (Word, Excel, etc.) via the Office Deployment Tool (ODT) without the "Click-to-Run" headache.
*   **System Debloat**: Surgical removal of Windows tracking, telemetry, and "Consumer Experience" adware.
*   **Activation Protocol**: Auto-detects and applies HWID/KMS activation for perpetual licensing (Educational/Research only).

---

## 💾 Core Modules & Features

### 1. 🏗️ Developer Environment (Visual Studio)
We automate the installer for **Visual Studio 2022 Enterprise**.
*   **Flags Used**: `--quiet`, `--norestart`, `--wait`.
*   **Workloads Installed**:
    *   `Microsoft.VisualStudio.Workload.ManagedDesktop` (C#/.NET)
    *   `Microsoft.VisualStudio.Workload.NetWeb` (ASP.NET)
    *   `Microsoft.VisualStudio.Workload.Python`
    *   `Microsoft.VisualStudio.Workload.Node`

### 2. 🧹 System Hygiene (Debloat)
Our script aggressively cleans the OS. **Warning**: This cannot be easily undone.
*   **Removes**: Cortana, Bing Weather, Get Help, Microsoft Tips, Solitaire, Mixed Reality Portal.
*   **Disables**: Telemetry Services, Advertising ID, Feedback Hub.
*   **Configures**: Dark Mode default, Show File Extensions, Taskbar alignment (Win 11).

### 3. 📦 Package Management (WinGet + Chocolatey)
Installs essential tools in bulk:
*   `vscode`, `git`, `docker-desktop`, `postman`
*   `7zip`, `vlc`, `spotify`
*   `powertoys`, `terminal`

---

## 🛠️ Usage Protocols

### 1. The "YOLO" Method (Recommended)
Just run the one-liner command above. The interactive TUI (Text User Interface) or GUI will guide you.

### 2. Manual Deployment
Clone the repository and run locally:

```powershell
git clone https://github.com/ACTechPRO/Microsoft-Ultimate-Installer.git
cd Microsoft-Ultimate-Installer
.\Microsoft Ultimate Installer.ps1
```

### 3. Uninstallation (The "Nuke" Option)
MUI includes a specialized diverse uninstaller to scrub Visual Studio and Office remnants if they become corrupted.

```powershell
.\Microsoft Ultimate Installer.ps1 -Uninstall
```
*   **Deep Clean**: Scans `C:\ProgramData` and Registry for leftover keys.
*   **Ghost Busters**: Scans Desktop and Start Menu for broken shortcuts and removes them.

---

## ❓ Troubleshooting (FAQ)

**Q: Why do I need Admin privileges?**
A: We modify system registry keys, install global packages, and remove system apps. This is impossible without elevated permissions.

**Q: Is this safe?**
A: Yes, if you are a developer. If you are a casual user who uses "Microsoft Tips" or "Cortana", do not run the Debloat module. The `irm` command downloads from our verified proxy directly to memory.

**Q: My Antivirus flagged this.**
A: Automating Windows settings often triggers false positives in Windows Defender (heuristics). You may need to add an exclusion or allow the script.

---

## ⚠️ Disclaimer

This tool offers "Activation" scripts (MAS/Ohook) strictly for **educational and research purposes**.
*   If you use this software in a corporate environment, **BUY A LICENSE**.
*   Microsoft, Windows, and Visual Studio are trademarks of Microsoft Corporation.
*   AC Tech Solutions is not affiliated with Microsoft.

---

<div align="center">
  
  **Executed by AC Tech Solutions**
  
  [🌐 Website](https://ac-tech.pro) • [🐙 GitHub](https://github.com/ACTechPRO)
  
  <sub>"Automation is the ultimate sophistication."</sub>

</div>
