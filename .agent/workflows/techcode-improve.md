# TechCode Improvement Workflow (Universal)

// turbo-all

## 1. Zero-Confirm Protocol (SafeToAutoRun: true)

> [!IMPORTANT]
> **AUTONOMY MODE: ON**
> - **Unified**: Adapts to project language.
> - **Safety**: Deletions backed up to `.gemini/trash` (Non-Recursive).
> - **Self-Healing**: Loops up to 10 times with stall detection.

## 2. Detection & Prep

```powershell
# 0. Safety First: Create Trash Directory
$trashBase = ".gemini/trash"
$trashPath = "$trashBase/$(Get-Date -Format 'yyyyMMdd_HHmm')"
if (-not (Test-Path $trashPath)) {
    New-Item -ItemType Directory -Force -Path $trashPath | Out-Null
}

function Safe-Delete {
    param($Path)
    if (Test-Path $Path) {
        # CRITICAL: Do not backup the trash folder into itself
        if ($Path -like "*$trashBase*") {
            Remove-Item $Path -Force -Recurse -ErrorAction SilentlyContinue
            return
        }
        
        Write-Host "Safe-Deleting $Path..." -ForegroundColor Yellow
        $parent = Split-Path $Path -Parent
        if ($parent) { $dest = Join-Path $trashPath $parent } else { $dest = $trashPath }
        if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Force -Path $dest | Out-Null }
        
        Copy-Item $Path $dest -Force -Recurse
        Remove-Item $Path -Force -Recurse
    }
}

function Safe-Update-File {
    param($Path, $NewContent)
    $tempPath = "$Path.tmp"
    [System.IO.File]::WriteAllText($tempPath, $NewContent, [System.Text.Encoding]::UTF8)
    
    # Language-Specific Safety Checks
    $errors = $null
    if ($Path -match "\.ps1$") {
        [void][System.Management.Automation.Language.Parser]::ParseFile($tempPath, [ref]$null, [ref]$errors)
    }
    
    if ($errors) {
        Write-Warning "Safe-Edit Failed: Syntax errors detected in fix for $Path"
        $errors | Format-Table -AutoSize
        Remove-Item $tempPath -Force
        return $false
    } else {
        Move-Item $tempPath $Path -Force
        Write-Host "Fixed (Verified): $Path" -ForegroundColor Green
        return $true
    }
}

# 1. Project Detection Strategy
$projectType = "Generic"
$validationCmd = $null
$fixCmd = $null

if (Test-Path pubspec.yaml) {
    Write-Host "Detected: FLUTTER Project" -ForegroundColor Cyan
    $projectType = "Flutter"
    $validationCmd = { flutter analyze 2>&1 }
    
    $fixCmd = { 
        Write-Host "Applying Flutter Fixes..."
        dart fix --apply
        dart format .
        flutter pub get
    }
    
    # Environment Prep
    if (Test-Path android/gradlew) { ./android/gradlew --stop 2>$null }
    flutter clean; flutter pub get
}
elseif ((Get-ChildItem -Filter *.ps1 -Recurse).Count -gt 0) {
    Write-Host "Detected: POWERSHELL Project" -ForegroundColor Cyan
    $projectType = "PowerShell"
    
    $validationCmd = { 
        if (Get-Command Invoke-ScriptAnalyzer -ErrorAction SilentlyContinue) {
            Invoke-ScriptAnalyzer -Path . -Recurse
        } else {
            Write-Warning "PSScriptAnalyzer not found."; $null 
        }
    }
    
    $fixCmd = {
        Write-Host "Applying PowerShell Fixes..."
        # 1. Trim Whitespace
        Get-ChildItem -Recurse -Filter *.ps1 | ForEach-Object {
            $c = Get-Content $_.FullName -Raw
            if ($c -match "(?m)\s+$") {
                $f = $c -replace "(?m)\s+$", ""
                Safe-Update-File -Path $_.FullName -NewContent $f
            }
        }
    }
}
```

## 3. The Self-Healing Loop

**Logic**:
1. Run `$validationCmd`
2. Analyze output: If clean -> Success.
3. Stall Detector: If error count matches previous 3 runs -> Abort (prevent infinite loop).
4. If limit reached -> Fail.

```powershell
$maxRetries = 10
$attempt = 1
$clean = $false
$prevCount = -1
$stallCounter = 0

if ($validationCmd) {
    do {
        Write-Host "=== Cycle $attempt/$maxRetries ($projectType) ===" -ForegroundColor Cyan
        
        # A. Verify
        $output = & $validationCmd
        
        # B. Analyze Issues
        $issuesFound = $false
        $count = 0
        
        if ($projectType -eq "Flutter") {
            $errs = $output | Where-Object { $_ -match "error •" }
            $count = ($errs | Measure-Object).Count
            if ($count -gt 0) { $issuesFound = $true }
        }
        elseif ($projectType -eq "PowerShell") {
            $count = ($output | Measure-Object).Count
            if ($output) { $issuesFound = $true }
        }
        
        Write-Host "Issues Found: $count" -ForegroundColor Gray
        
        if (-not $issuesFound) {
            Write-Host "✅ Project is CLEAN." -ForegroundColor Green
            $clean = $true
            break
        }
        
        # C. Stall Detection
        if ($count -eq $prevCount) {
            $stallCounter++
            if ($stallCounter -ge 3) {
                Write-Warning "🛑 Fix loop stalled (Issue count constant for 3 cycles). Stopping."
                break
            }
        } else {
            $stallCounter = 0
        }
        $prevCount = $count
        
        # D. Execute Fix
        if ($attempt -lt $maxRetries) {
            Write-Host "Applying Auto-Fixes..." -ForegroundColor Yellow
            & $fixCmd
        } else {
            Write-Error "❌ Critical issues remain after $maxRetries attempts."
        }
        
        $attempt++
    } while ($attempt -le $maxRetries)
} else {
    Write-Warning "No validation strategy for this project type."
}
```

## 4. Final Polish (Git)

```powershell
# Cleanup Trash (No recursion check needed here because we use Remove-Item direct)
Remove-Item ".gemini/trash" -Recurse -Force -ErrorAction SilentlyContinue

# Commit if clean
if ($clean) {
    git add .
    git commit -am "✨ TechCode Auto-Improvement ($projectType Optimized)"
}
```
