# VBA Project Scaffolding Tool - Complete Documentation

A comprehensive guide to using the `vba-tool` CLI for automated VBA project creation and management.

---

## Table of Contents

1. [Overview](#overview)
2. [Installation & Setup](#installation--setup)
3. [Command Reference](#command-reference)
4. [Templates](#templates)
5. [LocalUtility.bas Macros](#localutilitybas-macros)
6. [Development Workflow](#development-workflow)
7. [Configuration](#configuration)
8. [Troubleshooting](#troubleshooting)
9. [Advanced Usage](#advanced-usage)
10. [Examples](#examples)

---

## Overview

### What is vba-tool?

`vba-tool` is a command-line interface for rapidly creating VBA projects with seamless WSL + Windows Excel integration. It automates:

- **Directory structure creation** (WSL side for coding)
- **Excel file creation** (.xlsm files on Windows side)
- **VBA module import automation** (LocalUtility.bas pre-configured)
- **Git initialization** (optional)
- **Template-based project scaffolding**

### Key Benefits

✅ **One-command project creation** - No manual setup
✅ **Auto-configured paths** - Works immediately without tweaking
✅ **Fast iteration** - Hot-reload VBA modules without restarting Excel
✅ **Version control ready** - Git integration built-in
✅ **Flexible templates** - Simple, Standard, or Complex structures

### How It Works

```
┌─────────────────────────────────────────────────────────┐
│  vba-tool new MyProject                                 │
└─────────────────┬───────────────────────────────────────┘
                  │
         ┌────────┴────────┐
         │                 │
         ▼                 ▼
┌─────────────────┐  ┌──────────────────────────┐
│  WSL Side       │  │  Windows Side            │
│  (Code)         │  │  (Excel)                 │
├─────────────────┤  ├──────────────────────────┤
│ /home/.../vba/  │  │ C:\Users\Luis\Documents\ │
│ MyProject/      │  │ VBA_Projects\MyProject\  │
│ ├─ LocalUtil... │  │ └─ MyProject.xlsm        │
│ ├─ BaseSheets/  │  │    (with LocalUtility    │
│ └─ WorkbookOps/ │  │     already imported)    │
└─────────────────┘  └──────────────────────────┘
```

---

## Installation & Setup

### Prerequisites

**Required:**
- WSL2 with Ubuntu
- Windows with Excel installed
- Python 3.x
- `jq` command-line JSON processor

**Check if jq is installed:**
```bash
which jq
```

**Install jq if needed:**
```bash
sudo apt install jq
```

### First-Time Setup

**1. The tool is already installed at:**
```
/home/lrev47/dev/active/vba/.vba-tools/
```

**2. Alias is configured in `~/.zshrc`:**
```bash
alias vba-tool='/home/lrev47/dev/active/vba/.vba-tools/vba-tool'
```

**3. Load the alias (or restart terminal):**
```bash
source ~/.zshrc
```

**4. Verify installation:**
```bash
vba-tool help
vba-tool templates
```

**5. Configuration is already set up:**
- WSL code location: `/home/lrev47/dev/active/vba`
- Windows Excel location: `C:\Users\Luis\Documents\VBA_Projects`

### Excel Configuration (One-Time)

For VBA module import/reload to work, enable VBA project access:

1. Open Excel
2. **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. **Macro Settings**
4. ☑ Check **"Trust access to the VBA project object model"**
5. Click **OK**

⚠️ **Important:** Without this setting, LocalUtility.bas macros won't work!

---

## Command Reference

### `vba-tool help`

Display usage information and available commands.

```bash
vba-tool help
```

**Output:**
```
VBA Project Scaffolding Tool

USAGE:
    vba-tool <command> [options]

COMMANDS:
    init                        Initialize configuration
    new <project> [options]     Create new VBA project
    add-workbook <project> <workbook>  Add workbook to existing project
    templates                   List available templates
    help                        Show this help message
```

---

### `vba-tool init`

Initialize or reconfigure the tool. Prompts for configuration values.

```bash
vba-tool init
```

**Interactive prompts:**
- WSL base path (default: `/home/lrev47/dev/active/vba`)
- Windows base path (default: `C:\VBA_Projects`)
- Default template (simple/standard/complex)
- Auto-initialize git? (Y/n)

**When to use:**
- First time setup (already done for you)
- Changing default paths
- Switching default template
- Updating configuration

**Example:**
```bash
vba-tool init
# Follow prompts to customize settings
```

---

### `vba-tool new`

Create a new VBA project.

**Syntax:**
```bash
vba-tool new <project-name> [options]
```

**Required:**
- `<project-name>` - Name of the project (no spaces, use CamelCase or snake_case)

**Options:**

| Option | Description | Default |
|--------|-------------|---------|
| `--template <name>` | Template to use (simple/standard/complex) | standard |
| `--multi` | Create as multi-workbook project | single workbook |
| `--no-git` | Don't initialize git repository | git enabled |
| `--windows-path <path>` | Custom Windows path for Excel file | Documents\VBA_Projects |

**Examples:**

**Basic project (standard template):**
```bash
vba-tool new MyProject
```

**Simple template:**
```bash
vba-tool new QuickTest --template simple
```

**Complex template:**
```bash
vba-tool new EnterpriseApp --template complex
```

**Multi-workbook project:**
```bash
vba-tool new SalesSystem --multi
# Prompts for first workbook name
```

**No git initialization:**
```bash
vba-tool new TempProject --no-git
```

**Custom Windows path:**
```bash
vba-tool new SpecialProject --windows-path "D:\Projects\VBA"
```

**What happens when you run `vba-tool new MyProject`:**

1. Creates WSL directory structure:
   ```
   /home/lrev47/dev/active/vba/MyProject/
   └── MyProject/
       ├── LocalUtility.bas (auto-configured)
       ├── BaseSheets/
       │   └── Builders/
       └── WorkbookOperations/
           ├── Directives/
           └── Overview/
   ```

2. Creates Windows Excel file:
   ```
   C:\Users\Luis\Documents\VBA_Projects\MyProject\
   └── MyProject.xlsm (with LocalUtility.bas imported)
   ```

3. Initializes git repository (unless `--no-git`)

4. Creates `.vba-project.json` metadata file

---

### `vba-tool add-workbook`

Add another workbook to an existing project.

**Syntax:**
```bash
vba-tool add-workbook <project-name> <workbook-name>
```

**Required:**
- `<project-name>` - Existing project name
- `<workbook-name>` - Name for new workbook

**Example:**
```bash
# Create multi-workbook project
vba-tool new SalesSystem --multi
# Enter first workbook name: Invoices

# Add more workbooks
vba-tool add-workbook SalesSystem Reports
vba-tool add-workbook SalesSystem Dashboard
```

**Result:**
```
/home/lrev47/dev/active/vba/SalesSystem/
├── Invoices/
│   └── LocalUtility.bas
├── Reports/
│   └── LocalUtility.bas
└── Dashboard/
    └── LocalUtility.bas

C:\Users\Luis\Documents\VBA_Projects\SalesSystem\
├── Invoices.xlsm
├── Reports.xlsm
└── Dashboard.xlsm
```

---

### `vba-tool templates`

List all available templates with descriptions.

```bash
vba-tool templates
```

**Output:**
```
ℹ Available templates:

complex
  Advanced structure with testing, utilities, and configuration support.

simple
  Minimal structure for quick prototypes. Just LocalUtility.bas for rapid iteration.

standard
  Standard structure with BaseSheets and WorkbookOperations. Based on UsageFlow pattern.
```

---

## Templates

### Simple Template

**Best for:** Quick prototypes, learning VBA, single-file projects

**Structure:**
```
ProjectName/
└── LocalUtility.bas
```

**Use case:**
- Rapid experimentation
- Single-purpose macros
- Testing ideas
- Learning VBA development workflow

**Create with:**
```bash
vba-tool new QuickTest --template simple
```

---

### Standard Template

**Best for:** Most production projects, typical Excel automation

**Structure:**
```
ProjectName/
├── LocalUtility.bas
├── BaseSheets/
│   └── Builders/
│       └── (place sheet builder modules here)
└── WorkbookOperations/
    ├── Directives/
    │   └── (place directive modules here)
    └── Overview/
        └── (place overview modules here)
```

**Based on:** Your UsageFlow project pattern

**Use case:**
- Excel workbooks with multiple sheets
- Business logic separated into directives
- Sheet-specific operations
- Standard business applications

**Folder purposes:**
- **BaseSheets/** - Sheet-specific code (Sheet_NewUsage.bas, etc.)
- **BaseSheets/Builders/** - Modules that build/rebuild sheets
- **WorkbookOperations/Directives/** - Business logic modules
- **WorkbookOperations/Overview/** - Summary/reporting modules

**Create with:**
```bash
vba-tool new MyProject --template standard
# or just:
vba-tool new MyProject  # standard is default
```

**LocalUtility.bas features:**
- `ImportAllModules()` - Imports from BaseSheets + WorkbookOperations
- `ReloadAllModules()` - Removes and re-imports all modules
- `ReloadDevModules()` - Fast reload of Directive/Overview only

---

### Complex Template

**Best for:** Large projects, team development, projects requiring testing

**Structure:**
```
ProjectName/
├── LocalUtility.bas
├── BaseSheets/
│   └── Builders/
├── WorkbookOperations/
│   ├── Directives/
│   └── Overview/
├── Utilities/
│   └── (helper functions, shared code)
├── Tests/
│   └── (test modules: Test_*.bas)
└── Config/
    └── (configuration modules)
```

**Additional features:**
- **Utilities/** - Shared helper functions
- **Tests/** - Unit test modules
- **Config/** - Configuration management

**Use case:**
- Large-scale projects
- Multiple developers
- Requires testing
- Complex business logic
- Shared utility functions

**LocalUtility.bas features:**
- All standard template features
- `RunAllTests()` - Executes all Test_* modules
- Imports from all 5 folders

**Create with:**
```bash
vba-tool new EnterpriseApp --template complex
```

---

## LocalUtility.bas Macros

Every project includes a `LocalUtility.bas` module with pre-configured macros for VBA development.

### Available Macros

#### `ImportAllModules`

Import all `.bas` files from project folders into Excel.

**When to use:**
- First time opening the Excel file
- After cloning project from git
- After adding new .bas files

**How to run:**
1. Open Excel file
2. Press `Alt+F11` to open VBA Editor
3. Find `LocalUtility` module
4. Press `F5` or click **Run** → **Run Sub/UserForm**
5. Select `ImportAllModules`

**Or use the Immediate Window:**
```vba
ImportAllModules
```

**What it does:**
- Scans project folders for .bas files
- Skips modules that already exist
- Imports new modules
- Reports results to Debug window

**Output:**
```
=== ImportAllModules Started ===
Base Path: \\wsl.localhost\Ubuntu\home\lrev47\dev\active\vba\MyProject\MyProject\
Found 15 .bas files to import
  IMPORTED: Sheet_NewUsage
  IMPORTED: Directive_NewUsage
  SKIPPED (already exists): LocalUtility
...
=== ImportAllModules Complete ===
Total Files: 15
Imported: 14
Skipped: 1
Errors: 0
Time: 0.85 seconds
```

---

#### `ReloadAllModules`

Remove all modules (except LocalUtility) and re-import them from disk.

**When to use:**
- After making code changes in VS Code
- Want to pick up ALL changes
- Infrastructure modules changed

**⚠️ Note:** Slower than `ReloadDevModules` but more thorough.

**How to run:**
```vba
ReloadAllModules
```

**What it does:**
1. Removes all standard modules (except LocalUtility)
2. Re-imports all .bas files from disk
3. Reports results

**Typical output:**
```
=== ReloadAllModules Started ===
Step 1: Removing all existing modules...
  REMOVED: Sheet_NewUsage
  REMOVED: Directive_NewUsage
  ... (all modules removed)
Removed 14 modules

Step 2: Re-importing all modules from disk...
  IMPORTED: Sheet_NewUsage
  IMPORTED: Directive_NewUsage
  ... (all modules re-imported)
=== ReloadAllModules Complete ===
Total time: 2.15 seconds
```

---

#### `ReloadDevModules` ⚡ (Standard/Complex templates only)

Fast reload of only Directive and Overview modules.

**When to use:**
- Rapid development iteration
- Only changed Directive_* or Overview_* files
- Want fastest reload time

**How to run:**
```vba
ReloadDevModules
```

**What it does:**
1. Removes only Directive_* and Overview_* modules
2. Re-imports only from Directives/ and Overview/ folders
3. Keeps infrastructure modules loaded (faster!)

**Typical output:**
```
=== ReloadDevModules Started ===
Step 1: Removing Directive and Overview modules...
  REMOVED: Directive_NewUsage
  REMOVED: Overview_NewUsage
Removed 6 modules

Step 2: Re-importing from Directives/ and Overview/ folders...
Found 6 .bas files to import
  IMPORTED: Directive_NewUsage
  IMPORTED: Overview_NewUsage
=== ReloadDevModules Complete ===
Time: 0.45 seconds
```

**⚡ Speed comparison:**
- `ReloadDevModules`: ~0.5 seconds (fast!)
- `ReloadAllModules`: ~2.0 seconds (thorough)

---

#### `RunAllTests` (Complex template only)

Run all test modules (Test_*.bas).

**When to use:**
- After making code changes
- Before committing to git
- Regression testing

**How to run:**
```vba
RunAllTests
```

**What it does:**
- Finds all modules starting with `Test_`
- Runs `RunTests` sub in each module
- Reports pass/fail results

**Example output:**
```
=== Running All Tests ===
Running: Test_Calculations
  PASSED
Running: Test_DataValidation
  PASSED
Running: Test_Import
  FAILED: Expected 10, got 9
=== Tests Complete ===
Tests Run: 3
Passed: 2
Failed: 1
```

---

## Development Workflow

### Complete Workflow Example

**1. Create project:**
```bash
cd /home/lrev47/dev/active/vba
vba-tool new BillAndBudget --template standard
```

**2. Navigate to project:**
```bash
cd BillAndBudget/BillAndBudget
```

**3. Create your VBA modules:**
```bash
# Create a sheet module
code BaseSheets/Sheet_Budget.bas

# Create a directive
code WorkbookOperations/Directives/Directive_CalculateBudget.bas

# Create an overview
code WorkbookOperations/Overview/Overview_MonthlySummary.bas
```

**4. Open Excel (Windows):**
- Navigate to: `C:\Users\Luis\Documents\VBA_Projects\BillAndBudget\`
- Open `BillAndBudget.xlsm`

**5. Import modules (first time):**
```vba
' In Excel VBA Editor (Alt+F11)
' Run this once:
ImportAllModules
```

**6. Development iteration:**

**In VS Code (WSL):**
```bash
# Edit your directive
code WorkbookOperations/Directives/Directive_CalculateBudget.bas
# Make changes, save
```

**In Excel:**
```vba
' Reload just the directives (fast!)
ReloadDevModules

' Test your changes
' ... run your macro ...
```

**Repeat step 6** as many times as needed.

**7. Commit when ready:**
```bash
git add .
git commit -m "Add budget calculation feature"
```

---

### Fast Iteration Pattern

For maximum development speed:

```
┌──────────────────────────────────────────────┐
│  Edit Directive/Overview in VS Code          │
│  (make changes, save)                        │
└────────────────┬─────────────────────────────┘
                 │
                 ▼
┌──────────────────────────────────────────────┐
│  In Excel: ReloadDevModules   ← 0.5 sec!    │
└────────────────┬─────────────────────────────┘
                 │
                 ▼
┌──────────────────────────────────────────────┐
│  Test changes immediately                     │
└────────────────┬─────────────────────────────┘
                 │
                 ▼
                Works? ───Yes──→  Commit
                 │
                 No
                 │
                 └─────→ Back to step 1
```

This cycle can be as fast as **30 seconds** per iteration!

---

## Configuration

### Configuration File Location

```
/home/lrev47/dev/active/vba/.vba-tools/config.json
```

### Current Configuration

```json
{
  "wsl_base_path": "/home/lrev47/dev/active/vba",
  "windows_base_path": "C:\\Users\\Luis\\Documents\\VBA_Projects",
  "windows_base_path_wsl": "/mnt/c/Users/Luis/Documents/VBA_Projects",
  "default_template": "standard",
  "git_auto_init": true,
  "default_author": "lrev47"
}
```

### Configuration Options

| Option | Description | Default |
|--------|-------------|---------|
| `wsl_base_path` | Where VBA code is stored (WSL) | `/home/lrev47/dev/active/vba` |
| `windows_base_path` | Where Excel files are created (Windows) | `C:\Users\Luis\Documents\VBA_Projects` |
| `windows_base_path_wsl` | WSL mount point for Windows path | `/mnt/c/Users/Luis/Documents/VBA_Projects` |
| `default_template` | Template used if not specified | `standard` |
| `git_auto_init` | Auto-initialize git repos | `true` |
| `default_author` | Git author name | `lrev47` |

### Changing Configuration

**Method 1: Interactive (recommended)**
```bash
vba-tool init
# Follow prompts
```

**Method 2: Manual edit**
```bash
code /home/lrev47/dev/active/vba/.vba-tools/config.json
# Edit values directly
```

---

## Troubleshooting

### "vba-tool: command not found"

**Cause:** Alias not loaded in current shell.

**Solutions:**

**Option 1: Reload shell config**
```bash
source ~/.zshrc
```

**Option 2: Restart terminal**
```bash
exit
# Open new terminal
```

**Option 3: Use full path**
```bash
/home/lrev47/dev/active/vba/.vba-tools/vba-tool help
```

**Verify alias:**
```bash
alias | grep vba-tool
# Should show: vba-tool=/home/lrev47/dev/active/vba/.vba-tools/vba-tool
```

---

### "Could not import VBA module"

**Cause:** VBA project access not enabled in Excel.

**Solution:**

1. Open Excel
2. **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. **Macro Settings**
4. ☑ **"Trust access to the VBA project object model"**
5. Click **OK**, restart Excel

**Verify setting:**
- Try running `ImportAllModules` again
- Should work without errors

---

### Excel file created but empty

**Cause:** PowerShell script failed to import LocalUtility.bas.

**Solution:**

**Manually import LocalUtility.bas:**
1. Open the Excel file
2. Press `Alt+F11` (VBA Editor)
3. **File** → **Import File**
4. Navigate to WSL path: `\\wsl.localhost\Ubuntu\home\lrev47\dev\active\vba\ProjectName\WorkbookName\`
5. Select `LocalUtility.bas`
6. Click **Open**

**Then run:**
```vba
ImportAllModules
```

---

### "Folder not found" errors in LocalUtility

**Cause:** LocalUtility.bas has wrong path or folders don't exist.

**Solution 1: Verify path in LocalUtility.bas**
```vba
' Check this constant in LocalUtility.bas:
Const BASE_PATH As String = "\\wsl.localhost\Ubuntu\home\lrev47\dev\active\vba\ProjectName\WorkbookName\"
```

**Solution 2: Create missing folders**
```bash
cd /home/lrev47/dev/active/vba/ProjectName/WorkbookName
mkdir -p BaseSheets/Builders
mkdir -p WorkbookOperations/Directives
mkdir -p WorkbookOperations/Overview
```

---

### Changes not reflected after ReloadDevModules

**Cause:** Changed infrastructure modules, not directive/overview modules.

**Solution:**

Use `ReloadAllModules` instead:
```vba
ReloadAllModules
```

**Or:** Restart Excel to force full reload.

---

### Git repository issues

**Problem:** Git not initialized or wrong remote.

**Solutions:**

**Initialize git manually:**
```bash
cd /home/lrev47/dev/active/vba/ProjectName
git init
```

**Check git status:**
```bash
git status
```

**Add remote:**
```bash
git remote add origin https://github.com/username/repo.git
```

---

## Advanced Usage

### Creating Custom Templates

**1. Create template directory:**
```bash
mkdir -p /home/lrev47/dev/active/vba/.vba-tools/templates/mytemplate
```

**2. Create template.json:**
```json
{
  "name": "mytemplate",
  "description": "My custom template for special projects",
  "folders": [
    "Modules",
    "Classes",
    "Utilities",
    "Data"
  ]
}
```

**3. Create LocalUtility.bas template:**

Use placeholders:
- `{{WSL_WORKBOOK_PATH}}` - WSL path to workbook
- `{{WINDOWS_WORKBOOK_PATH}}` - Windows path to workbook
- `{{WORKBOOK_NAME}}` - Workbook name

**4. Use custom template:**
```bash
vba-tool new MyProject --template mytemplate
```

---

### Multi-Environment Setup

**Development, Staging, Production:**

**Create separate config files:**
```bash
cp config.json config.dev.json
cp config.json config.prod.json
```

**Edit each:**
```json
// config.dev.json
{
  "windows_base_path": "C:\\Users\\Luis\\Documents\\VBA_Projects_Dev",
  ...
}

// config.prod.json
{
  "windows_base_path": "C:\\Users\\Luis\\Documents\\VBA_Projects_Prod",
  ...
}
```

**Switch environments:**
```bash
# Use dev
cp .vba-tools/config.dev.json .vba-tools/config.json

# Use prod
cp .vba-tools/config.prod.json .vba-tools/config.json
```

---

### Batch Project Creation

**Create multiple projects:**
```bash
#!/bin/bash
projects=(
  "Invoicing"
  "Inventory"
  "Reporting"
  "Dashboard"
)

for project in "${projects[@]}"; do
  vba-tool new "$project" --template standard
done
```

---

### Integration with Git Workflows

**Pre-commit hook example:**

Create `.git/hooks/pre-commit`:
```bash
#!/bin/bash
# Ensure all .bas files are properly formatted
find . -name "*.bas" -type f -exec dos2unix {} \;
```

**Make executable:**
```bash
chmod +x .git/hooks/pre-commit
```

---

## Examples

### Example 1: Quick Prototype

**Goal:** Test a new calculation algorithm quickly.

```bash
# Create simple project
vba-tool new CalcTest --template simple

# Navigate to folder
cd CalcTest/CalcTest

# Create calculation module
code CalculationEngine.bas
```

**In CalculationEngine.bas:**
```vba
Attribute VB_Name = "CalculationEngine"
Option Explicit

Public Sub TestCalculation()
    Debug.Print "Result: " & Calculate(10, 5)
End Sub

Private Function Calculate(a As Double, b As Double) As Double
    Calculate = a * b + (a / b)
End Function
```

**In Excel:**
```vba
ImportAllModules
TestCalculation
```

**Iterate:**
- Edit in VS Code
- Run `ReloadAllModules` in Excel
- Test again

**Total time:** ~2 minutes from idea to working prototype!

---

### Example 2: Production Budget Tracker

**Goal:** Create a multi-sheet budget tracker with calculations.

```bash
# Create standard project
vba-tool new BudgetTracker --template standard

cd BudgetTracker/BudgetTracker
```

**Create modules:**

```bash
# Sheet builders
code BaseSheets/Sheet_Budget.bas
code BaseSheets/Sheet_Expenses.bas
code BaseSheets/Sheet_Summary.bas
code BaseSheets/Builders/Module_BuildAll.bas

# Business logic
code WorkbookOperations/Directives/Directive_CalculateTotals.bas
code WorkbookOperations/Directives/Directive_ValidateEntries.bas

# Reports
code WorkbookOperations/Overview/Overview_MonthlySummary.bas
```

**Development workflow:**
1. Open `BudgetTracker.xlsm`
2. Run `ImportAllModules`
3. Edit directives in VS Code
4. Run `ReloadDevModules` (fast!)
5. Test functionality
6. Repeat 3-5 until complete
7. Commit to git

---

### Example 3: Enterprise Multi-Workbook System

**Goal:** Sales system with multiple interconnected workbooks.

```bash
# Create multi-workbook project
vba-tool new SalesSystem --template complex --multi
# Enter first workbook: Invoices

# Add more workbooks
vba-tool add-workbook SalesSystem Customers
vba-tool add-workbook SalesSystem Reports
vba-tool add-workbook SalesSystem Dashboard
```

**Result:**
```
SalesSystem/
├── Invoices/
│   ├── BaseSheets/
│   ├── WorkbookOperations/
│   ├── Utilities/
│   ├── Tests/
│   └── Config/
├── Customers/
│   └── (same structure)
├── Reports/
│   └── (same structure)
└── Dashboard/
    └── (same structure)
```

**Share code between workbooks:**
```bash
# Create shared utilities
mkdir -p SalesSystem/Shared
code SalesSystem/Shared/CommonUtilities.bas

# Symlink or copy to each workbook
cp SalesSystem/Shared/CommonUtilities.bas SalesSystem/Invoices/Utilities/
cp SalesSystem/Shared/CommonUtilities.bas SalesSystem/Customers/Utilities/
```

---

### Example 4: Test-Driven Development

**Goal:** Build tested, reliable code.

```bash
# Create project with testing support
vba-tool new DataProcessor --template complex

cd DataProcessor/DataProcessor
```

**Create test module:**
```bash
code Tests/Test_DataValidation.bas
```

**In Test_DataValidation.bas:**
```vba
Attribute VB_Name = "Test_DataValidation"
Option Explicit

Public Sub RunTests()
    Debug.Print "Running DataValidation tests..."

    ' Test 1: Valid data
    If ValidateData("12345") Then
        Debug.Print "  PASS: Valid data accepted"
    Else
        Debug.Print "  FAIL: Valid data rejected"
    End If

    ' Test 2: Invalid data
    If Not ValidateData("abc") Then
        Debug.Print "  PASS: Invalid data rejected"
    Else
        Debug.Print "  FAIL: Invalid data accepted"
    End If
End Sub
```

**Create implementation:**
```bash
code Utilities/DataValidation.bas
```

**Test workflow:**
1. Open Excel
2. Run `ImportAllModules`
3. Run `RunAllTests`
4. See results
5. Fix failing tests
6. Run `ReloadAllModules`
7. Run `RunAllTests` again
8. Repeat until all pass

---

## Quick Reference Card

### Most Common Commands

```bash
# List templates
vba-tool templates

# Create new project (standard)
vba-tool new MyProject

# Create simple project
vba-tool new MyProject --template simple

# Create complex project
vba-tool new MyProject --template complex

# Add workbook to project
vba-tool add-workbook MyProject SecondBook

# Get help
vba-tool help
```

### Most Common VBA Macros

```vba
' First time setup
ImportAllModules

' Fast reload (directives/overview only)
ReloadDevModules

' Full reload (everything)
ReloadAllModules

' Run all tests (complex template)
RunAllTests
```

### File Locations

```
WSL Code:     /home/lrev47/dev/active/vba/ProjectName/
Windows Excel: C:\Users\Luis\Documents\VBA_Projects\ProjectName\
Tool:         /home/lrev47/dev/active/vba/.vba-tools/
Config:       /home/lrev47/dev/active/vba/.vba-tools/config.json
```

---

## Getting Help

**View this documentation:**
```bash
code /home/lrev47/dev/active/vba/.vba-tools/DOCUMENTATION.md
```

**View README:**
```bash
code /home/lrev47/dev/active/vba/.vba-tools/README.md
```

**View configuration:**
```bash
cat /home/lrev47/dev/active/vba/.vba-tools/config.json
```

**Check tool version:**
```bash
vba-tool help
```

---

## Summary

The `vba-tool` CLI provides a complete VBA development environment with:

✅ **Automated project creation** - One command setup
✅ **Three flexible templates** - Simple, Standard, Complex
✅ **Fast development iteration** - Sub-second reload times
✅ **WSL + Windows integration** - Seamless workflow
✅ **Git version control** - Built-in support
✅ **Test support** - TDD-ready (complex template)

**Start creating VBA projects in seconds instead of hours!**

```bash
vba-tool new MyFirstProject
```

Happy coding! 🚀
