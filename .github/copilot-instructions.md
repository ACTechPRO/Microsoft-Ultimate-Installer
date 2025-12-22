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
| `deepseek-r1:32b` | **Reasoning Engine** | Planning, architecture, deep thought (Expect Chain of Thought) |

**Configuration**: Models run locally via Ollama on the user's RTX 4080 (Hybrid GPU/RAM). Model directory: `D:\TechAI\TechAgents`.

### How to Use Local Models (MANDATORY for Agents)

Agents **MUST** use local models via `delegate.ps1`. Do NOT run `ollama` directly.

**Use Cases (prioritize local models for these):**
- **Coding (>30 lines)**: `delegate.ps1 -Type builder -Prompt "..."`
- **Planning**: `delegate.ps1 -Type planner -Prompt "..."`

> **IMPORTANT**: Use `32b` models for quality. The script handles system prompt injection automatically.

---

## Local AI Usage (The "Architect-Builder" Protocol)

You have access to a **Supercomputer (i9/4080/64GB)**.
- **Role**: You are the **ARCHITECT** (Cloud Agent).
- **Resource**: The Local AI is your **BUILDER** (Worker).

### Core Directive: "Architect First, Then Delegate"
Do not just "pass the buck". You must analyze, design, and provide structured context *before* delegating execution to the local model.

### ⚠️ MANDATORY CHECKPOINT (Before Any Work)

Before starting work on ANY planning or coding task, **STOP** and answer these questions:

| Condition | Action |
|-----------|--------|
| **Complex Planning** | **Joint Effort**: You outline high-level strategy (Architect), then delegate detailed execution plan to `deepseek-r1:32b` (Planner). |
| **Code > 30 lines** | **Delegate**: Create a "Context-Rich Prompt" and delegate to `qwen2.5-coder:32b`. |
| **Code < 30 lines** | **Execute**: Handle directly for speed and flow. |

> [!CAUTION]
> **Protocol Violation Alert**: If you find yourself writing >30 lines of code without having run a local model first, **STOP IMMEDIATELY**. You are violating the Architect-Builder Protocol. Delegate now.

### How to Delegate (MANDATORY SCRIPT)

**You MUST use `delegate.ps1`. Do NOT run `ollama` directly.**

**For Coding (>30 lines)**:
```powershell
D:\TechAI\TechAgents\delegate.ps1 -Type builder -Prompt "Context: [your context]. Task: [your task]"
```

**For Planning/Architecture**:
```powershell
D:\TechAI\TechAgents\delegate.ps1 -Type planner -Prompt "[your analysis request]"
```

**After Execution**:
- Script saves timestamped output to `D:\TechAI\TechAgents\Ollama\Outputs\`
- Script ALSO updates the latest pointer: `D:\TechAI\TechAgents\draft_output.txt`
- Read with: `view_file D:\TechAI\TechAgents\draft_output.txt`
- Script has per-model timeouts (builder: 10 min, planner: 15 min)

### ⏱️ Timeout Expectations (BE PATIENT)

| Model | Typical Time | Default Timeout |
|-------|--------------|-----------------|
| **DeepSeek (planner)** | 5-10 minutes | 15 min |
| **Qwen (builder)** | 1-3 minutes | 10 min |

- **THIS IS NORMAL** — Local models are slower but smarter
- Script shows progress every 30 seconds ("Still running...")
- After running `delegate.ps1`, use `command_status` with `WaitDurationSeconds: 600`
- Only consider failure if script exits with an error code

### ⛔ No Silent Fallback (CRITICAL)

If the local model times out or fails:
1. **RETRY** with `-TimeoutMinutes 20` for very complex tasks
2. **If still fails**, check if Ollama is running: `ollama ps`
3. **ONLY after retry fails**, proceed autonomously with a note: "Local model unavailable, proceeding manually."

**NEVER** silently do the work yourself after a single timeout. This defeats the protocol.

### Hybrid Prompt Architecture
The script enforces a **Fixed + Flexible** structure:
- **Fixed**: System Prompt (`builder.md`/`planner.md`) is always prepended automatically
- **Flexible**: You provide context and task via `-Prompt` argument
- **Never** add "You are a coding assistant" — the script handles this

### 📋 Evidence Requirement (MANDATORY)

Your response **MUST** show **PROOF** of local model usage:

✅ **Valid Evidence**:
- The exact `delegate.ps1` command you ran
- Confirmation you read `draft_output.txt`
- Quotes or summaries from local model output

❌ **Invalid (VIOLATIONS)**:
- "I delegated to the model" (no command shown)
- Presenting your own work without delegation
- Skipping delegation without explicit user permission

### The "Zero-Context" Workflow
NEVER paste heavy local output into the chat.
1.  **Delegate**: `delegate.ps1 -Type ... -Prompt "..."`
2.  **Review**: Read `D:\TechAI\TechAgents\draft_output.txt`.
3.  **Refine**: Edit if necessary.
4.  **Result**: You look like a genius, and the Cloud Context remains empty.

### ✅ Correct Workflow Example

**User Request**: "Create a Python service for PDF processing"

**Correct Agent Behavior**:
```powershell
# Step 1: Delegate architecture planning
D:\TechAI\TechAgents\delegate.ps1 -Type planner -Prompt "Design a Python service architecture for PDF processing with folder watching, OCR extraction, and SQLite storage."

# Step 2: Review and refine
view_file D:\TechAI\TechAgents\draft_output.txt

# Step 3: Delegate code generation
D:\TechAI\TechAgents\delegate.ps1 -Type builder -Prompt "Context: [paste refined plan]. Task: Write the Python watcher.py module."

# Step 4: Review and finalize
view_file D:\TechAI\TechAgents\draft_output.txt
```

**Incorrect Behavior**: Running `ollama run` directly or writing >30 lines yourself.

### 📋 Available Workflows

Use the **`/delegate`** workflow for any planning or coding task. The workflow provides step-by-step instructions with auto-execution enabled.

**Location**: `.agent/workflows/delegate.md` in each project


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

## Local AI (Architect-Builder Protocol)

- **Role**: You are the **ARCHITECT**. Local AI is the **BUILDER**.
- **Delegation**: Use `D:\TechAI\TechAgents\delegate.ps1` for all calls.
- **Coding (>30 lines)**: `delegate.ps1 -Type builder -Prompt "..."`
- **Planning**: `delegate.ps1 -Type planner -Prompt "..."`
- **Timeout**: Models take 2-5 minutes — **BE PATIENT**.

> [!IMPORTANT]
> Never run `ollama` directly. Use `delegate.ps1`. Never silently fall back after a timeout.


---

> **Auto-generated by sync-rules.ps1** on 2025-12-21 20:25:28
> Source: D:\GLOBAL_RULES.md + D:\Microsoft Ultimate Installer\PROJECT_RULES.md
> Do not edit this file directly. Edit source files and re-run sync.
