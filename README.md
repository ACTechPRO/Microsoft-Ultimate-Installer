<div align="center">

# Microsoft Ultimate Installer
### The Last Windows Setup Script You'll Ever Need.

![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![Version](https://img.shields.io/badge/Version-1.0.0-green?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-orange?style=for-the-badge)

</div>

---

## ⚡ Initiate Sequence (Admin)

Copy and paste this into a **PowerShell (Admin)** terminal. Watch the magic happen.

```powershell
irm ult.ac-tech.pro | iex
```

> **Note**: Requires Internet Connection. 

---

## 📡 Transmission: Overview

**Microsoft Ultimate Installer (MUI)** is a sophisticated automation framework designed for power users, developers, and sysadmins. It bypasses the bloat, automates the tedious, and delivers a pristine development environment in minutes.

### Why MUI?
*   **Zero-Click Visual Studio**: Installs Enterprise/Pro editions silently with your curated workloads.
*   **Office 365 Automation**: Deploys the full suite (Word, Excel, etc.) without the "Click-to-Run" headache.
*   **System Debloat**: Surgical removal of Windows tracking, telemetry, and "Consumer Experience" adware.
*   **Activation Protocol**: Auto-detects and applies HWID/KMS activation for perpetual licensing (Educational/Research only).

---

## 💾 Core Modules

| Module | Status | Description |
| :--- | :---: | :--- |
| **Visual Studio** | 🟢 | Silent install of IDE + Workloads (C#, C++, Python, Web). |
| **Microsoft 365** | 🟢 | Full Office suite deployment via ODT. |
| **WinGet Integration** | 🟢 | Bulk install of essential tools (`vscode`, `git`, `docker`, `spotify`). |
| **Debloater** | 🟡 | Removes Cortana, OneDrive (optional), and Telemetry. |
| **Activator** | 🔴 | *Hidden Module*: Legacy activation support. |

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

---

## 💬 Community & Support

*   [**Discussions**](https://github.com/ACTechPRO/Microsoft-Ultimate-Installer/discussions): Ask questions, request features, or share your config.
*   [**Issues**](https://github.com/ACTechPRO/Microsoft-Ultimate-Installer/issues): Report bugs or protocol failures.

---

<div align="center">
  
  **Executed by AC Tech Solutions**
  
  [🌐 Website](https://ac-tech.pro) • [🐙 GitHub](https://github.com/ACTechPRO)
  
  <sub>"Automation is the ultimate sophistication."</sub>

</div>
