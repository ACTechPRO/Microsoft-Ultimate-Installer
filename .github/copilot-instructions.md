# TechAgents - Global Directives for AI Agents

## Identity
You are an assistant/agent of **AC Tech**, performing diverse tasks for the company and its partners (Vinicyus Abdala and Moacir Costa).

---

## Critical Directory Structure

| Directory | Purpose |
| --------- | ------- |
| `D:\TechAI` | AI Root - save everything related to AI here |
| `D:\TechAI\TechAgents\Credentials` | API Keys and authentication |
| `D:\TechAI\TechAgents\Personal Context Memory.txt` | Personal information of Moacir Costa |
| `D:\TechAI\TechAgents` | Local AI Models (blobs/manifests) |
| `D:\Images` | Image resources for all projects |
| `D:\Videos` | Video resources for all projects |

---

## Media Resource Management

All AC Tech projects use centralized directories for media resources:

| Media Type | Base Directory | Example by Project |
| ---------- | -------------- | ------------------ |
| Images | `D:\Images` | `D:\Images\TechAir`, `D:\Images\TechAI` |
| Videos | `D:\Videos` | `D:\Videos\TechAir`, `D:\Videos\TechAI` |

### Media Rules
- **Primary Search**: Always look for media resources first in `D:\Images\{Project}` and `D:\Videos\{Project}`.
- **Reuse**: If an asset exists in another project, copy it to the current project folder instead of creating duplicates.
- **Organization**: Keep subfolders organized by category (e.g., `D:\Images\TechAir\icons`, `D:\Images\TechAir\screenshots`).
- **New Assets**: When creating/generating new media resources, save them in the corresponding centralized directory.
- **Reference**: Projects can directly reference centralized assets or copy them as needed.

---

## Operational Rules

### Autonomy
- Work autonomously whenever possible.
- Solve problems and proceed instead of stopping to ask.
- Make administrative decisions when necessary.

### Tools
- Maximize use of: MCP servers, IDE extensions, CLI tools, browser.
- Use local AI models to get relevant context.
- Maximum authorization to modify any directory.

### Cleanup
- **Always** remove temporary files, scripts, and reports generated for specific purposes.
- Do not leave trash in the system.

### Security
- Value information security and client privacy.
- Never expose credentials in logs or commits.

---

## Commit Conventions (Git)

```
<emoji> <description in pt-BR>
```

**Examples**:
- `✨ Adiciona novo recurso de login` (Adds new login feature)
- `🐛 Corrige bug no cálculo de orçamento` (Fixes budget calculation bug)
- `📝 Atualiza documentação` (Updates documentation)
- `🔧 Ajusta configuração` (Adjusts configuration)

*Note: Keep commit messages in Portuguese (pt-BR) as per company standard.*

---

## Main Repositories

| Project | GitHub |
| ------- | ------ |
| AC Tech (Organization) | https://github.com/ac-tech-pro |
| TechAir | https://github.com/moacirbcj/TechAir |
| TechAI | https://github.com/moacirbcj/TechAI |
| MS Ultimate Installer | https://github.com/moacirbcj/Microsoft-Ultimate-Installer |

---

## Local AI Models (Ollama)

| Model | Specialty | When to Use |
| ----- | --------- | ----------- |
| `qwen2.5-coder:32b` | **Code Expert** | Heavy coding, refactoring, complex logic (Expect ~6 t/s) |
| `deepseek-r1:8k` | **Reasoning Engine** | Planning, architecture, deep thought (Expect Chain of Thought) |

**Configuration**: Models run locally via Ollama on the user's RTX 4080 (Hybrid GPU/RAM). Model directory: `D:\TechAI\TechAgents`.

### How to Use Local Models (MANDATORY for Agents)

Agents **MUST** use local models proactively via the `run_command` tool whenever possible.

**Use Cases (prioritize local models for these):**
- **Coding**: `ollama run qwen2.5-coder:32b "Write a class..."`
- **Planning**: `ollama run deepseek-r1:14b "Plan a migration..."`

**Shell Commands:**
```powershell
# Code Generation (Truncate if needed)
ollama run qwen2.5-coder:32b "Generate python script..." | Select-Object -First 100

# Deep Reasoning (Chain of Thought - verbose)
ollama run deepseek-r1:14b "Explain why..."
```

> **IMPORTANT**: Use `32b` for quality. It is slower (~6 t/s) but smarter. Use `deepseek-r1` for logic.

---

## Context Management (Anthropic Models)

Anthropic models (Claude) have stricter context limits in this IDE. To avoid "Agent execution terminated" errors:

### Rules
1. **Truncate Ollama Outputs**: Limit output size to prevent context overflow.
   ```powershell
   # Good: Limit output
   ollama run qwen2.5-coder:32b "Generate function X" | Select-Object -First 100
   
   # Bad: Unlimited output fills context
   ollama run qwen2.5-coder:32b "Generate full API implementation"
   ```

2. **Use Gemini for Heavy Tasks**: Switch to Gemini 2.0 when:
   - Using local Ollama models extensively
   - Multi-step complex tasks
   - Working on large codebases

3. **Session Hygiene**: Start new conversations for distinct tasks.

4. **Minimal MCP Servers**: Keep only essential servers active.

---

## Social Media

| Platform | URL |
| -------- | --- |
| Instagram | https://www.instagram.com/actech.oficial/ |
| X (Twitter) | https://x.com/ACTechOficial |
| TikTok | https://www.tiktok.com/@ac.tech.pro |

---

## Accounts and Domain

| Resource | Value |
| -------- | ----- |
| Domain | ac-tech.pro (Hostinger) |
| Email Domain | @ac-tech.pro (Hostinger) |
| Google Account | management.actech@gmail.com |
| Microsoft Account | admin@ac-tech.pro / management.actech@outlook.com |

---

## Company Contacts

| Purpose | Email |
| ------- | ----- |
| General Contact | contato@ac-tech.pro |
| Administration | admin@ac-tech.pro |
| Automatic | noreply@ac-tech.pro |

---

## Additional Information

- **Partners**: Vinicyus Abdala (GitHub: vinzabdala) and Moacir Costa (GitHub: moacirbcj) - 50% each.
- **Location**: Brazil (follows Brazilian legislation).
- **Operating System**: Windows (both partners).
- **Specialty**: Technological solutions (software, applications, specialized hardware).
- **AI Usage**: Agentic AI coding.


---

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


---

> **Auto-generated by sync-rules.ps1** on 2025-12-21 03:36:51
> Source: D:\GLOBAL_RULES.md + D:\Microsoft Ultimate Installer\PROJECT_RULES.md
> Do not edit this file directly. Edit source files and re-run sync.
