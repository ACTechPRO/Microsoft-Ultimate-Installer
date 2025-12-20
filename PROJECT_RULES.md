# Microsoft Ultimate Installer - Project Specific Rules

## Project Information
- **Type**: Single PowerShell Script for installing/removing Microsoft products
- **Repository**: https://github.com/moacirbcj/Microsoft-Ultimate-Installer

## Tech Stack
- PowerShell (Single autonomous script)
- Windows API (WPF/XAML)
- Winget/Chocolatey (auto-managed)

## Media Resources
- **Images**: `D:\Images\Microsoft Ultimate Installer`
- **Videos**: `D:\Videos\Microsoft Ultimate Installer`

## Operational Rules

### Core Architecture
- **Single Autonomous File**: The script (`Microsoft Ultimate Installer.ps1`) must be 100% self-contained. No external module dependencies.
- **Portability**: Must run from ANY directory/computer. Path handling must be relative/dynamic.
- **Compatibility**: Support ALL PowerShell versions (5.1+, Core).
- **Privileges**: Auto-verify and request Administrator rights immediately.
- **Self-Dependency**: Use script to identify/download/install any required tools (e.g., Winget) automatically.

### Functionality
- **Installation**: Install multiple Microsoft apps, ALWAYS selecting the latest available version (including Insiders).
- **OS Management**: Allow changing Windows OS Version and Insiders Program channel.
- **Privacy & Optimization**: 
    - Offer options to maximize privacy (reduce data sent to MS) for provided apps.
    - Minimize background usage of services like Click-to-Run.
- **Modes**:
    - **Express Install**: Pre-selected software + Windows version modification.
    - **Complete Uninstall**: Forcibly remove ALL software offered by the script (deep clean: registries, folders) *EXCEPT* Visual Studio Code and Winget.
- **Licensing**: Handle licensing for both Windows and installed Apps.

### User Interface (WPF)
- **Window Behavior**:
    - Draggable by clicking *anywhere* (except buttons).
    - **Transparency**: Windows become transparent when dragged (except footer).
    - **Controls**: Always functional Minimize, Maximize, Close buttons (Top-Right).
- **Branding (Base64 Embeds)**:
    - **Header**: Use `D:\Images\Microsoft Ultimate Installer\Microsoft Ultimate Installer.png` (High Res).
    - **Footer**: Use `D:\Images\AC Tech\AC Tech Transparente Invertido.ico` (High Res/Layers).
    - Both logos must be present in ALL windows.
- **Progress**: Show descriptive, verbose information on current stage.
- **Stealth**: All background installations MUST be silent. Only the Script UI should be visible.
- **Clarity**: Buttons must have descriptive labels.

### System Integrity & Cleanup
- **Shortcuts**: Create Desktop shortcuts for ALL installed apps. Use proper icons.
- **Localization**: Auto-detect OS language and set script language accordingly.
- **Cleanup**: 
    - Self-clean after Success, Failure, or Cancel.
    - *EXCEPT*: Do NOT remove VS Code or Winget.
- **Logging**: 
    - Generate `Microsoft Ultimate Installer Log.txt` on Desktop **ONLY** if installation FAILS.
    - No logs for success/cancel.

### Implementation Guidelines
- **Code Quality**: Keep code organized, transparent, and human-readable.
- **WPF Integrity**: Ensure UI thread never hangs (use Runspaces/Jobs).
