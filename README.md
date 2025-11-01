# VBA Project Scaffolding Tool

A command-line tool for quickly creating VBA projects with WSL + Windows Excel integration.

## Quick Start

```bash
# First time setup
vba-tool init

# Create a new project
vba-tool new MyProject --template standard

# List available templates
vba-tool templates

# Add another workbook to existing project
vba-tool add-workbook MyProject SecondWorkbook
```

## Features

- **Automated Project Creation** - One command creates both WSL directory structure and Windows Excel files
- **Template System** - Three built-in templates (simple, standard, complex) plus ability to create custom ones
- **Path Auto-Configuration** - LocalUtility.bas files are automatically configured with correct WSL/Windows paths
- **Git Integration** - Optional automatic git initialization with appropriate .gitignore
- **Multi-Workbook Support** - Easily add multiple workbooks to a single project

## Templates

### Simple
Minimal structure for quick prototypes. Just a LocalUtility.bas for rapid iteration.

**Use for:** Quick experiments, single-module projects, learning VBA

**Structure:**
```
ProjectName/
└── LocalUtility.bas
```

### Standard
Standard structure with BaseSheets and WorkbookOperations folders. Based on UsageFlow pattern.

**Use for:** Most production projects, typical Excel automation tasks

**Structure:**
```
ProjectName/
├── LocalUtility.bas
├── BaseSheets/
│   └── Builders/
└── WorkbookOperations/
    ├── Directives/
    └── Overview/
```

### Complex
Advanced structure with testing, utilities, and configuration support.

**Use for:** Large projects, team development, projects requiring testing

**Structure:**
```
ProjectName/
├── LocalUtility.bas
├── BaseSheets/
├── WorkbookOperations/
├── Utilities/
├── Tests/
└── Config/
```

## Workflow

Your typical development workflow:

1. **Create project** - `vba-tool new MyProject`
2. **Code in VS Code (WSL)** - Edit .bas files in your project directory
3. **Open Excel (Windows)** - Open the .xlsm file created in C:\VBA_Projects\MyProject\
4. **Run macro** - In Excel, run `LocalUtility.ImportAllModules` or `LocalUtility.ReloadAllModules`
5. **Test** - Test your functionality
6. **Iterate** - Repeat steps 2-5

### Fast Iteration

For rapid development iteration:

```vba
' First time: Import all modules
ImportAllModules

' During development: Just reload changed modules
ReloadDevModules  ' Only reloads Directive/Overview modules (fast!)

' Or reload everything
ReloadAllModules  ' Reloads all modules (slower but complete)
```

## Configuration

Configuration is stored in `config.json`:

```json
{
  "wsl_base_path": "/home/lrev47/dev/active/vba",
  "windows_base_path": "C:\\VBA_Projects",
  "windows_base_path_wsl": "/mnt/c/VBA_Projects",
  "default_template": "standard",
  "git_auto_init": true,
  "default_author": "lrev47"
}
```

Edit this file to customize your setup or run `vba-tool init` again.

## Commands Reference

### `vba-tool init`
Initialize or reconfigure the tool. Asks for paths and preferences.

### `vba-tool new <project> [options]`
Create a new VBA project.

**Options:**
- `--template <name>` - Use specified template (simple/standard/complex)
- `--multi` - Create as multi-workbook project
- `--no-git` - Don't initialize git repository
- `--windows-path <path>` - Custom Windows path for Excel file

**Examples:**
```bash
vba-tool new MyProject
vba-tool new BillAndBudget --template simple
vba-tool new SalesFlow --template standard --multi
```

### `vba-tool add-workbook <project> <workbook>`
Add another workbook to an existing project.

**Example:**
```bash
vba-tool add-workbook UsageFlow UsageReporting
```

### `vba-tool templates`
List all available templates with descriptions.

## Excel Setup Requirements

For the LocalUtility.bas import/reload macros to work, you need to enable VBA project access:

1. Open Excel
2. File > Options > Trust Center > Trust Center Settings
3. Macro Settings
4. Check "Trust access to the VBA project object model"

This only needs to be done once per Excel installation.

## Directory Structure

```
/home/lrev47/dev/active/vba/
├── .vba-tools/              ← The tool itself
│   ├── vba-tool            ← Main CLI script
│   ├── config.json          ← Configuration
│   ├── lib/                 ← Helper scripts
│   │   ├── config-manager.sh
│   │   ├── template-engine.sh
│   │   └── create-excel.py
│   └── templates/           ← Project templates
│       ├── simple/
│       ├── standard/
│       └── complex/
│
├── UsageFlow/               ← Your projects
│   ├── UsageWorkbook/
│   ├── UsageTracker/
│   └── CreateUsageWorkbook/
│
└── BillAndBudget/           ← Future projects
```

## Troubleshooting

### "vba-tool: command not found"
The alias may not be loaded in your current shell. Either:
- Restart your terminal, or
- Run: `source ~/.bashrc`, or
- Use the full path: `/home/lrev47/dev/active/vba/.vba-tools/vba-tool`

### "Could not import VBA module"
Make sure you've enabled "Trust access to VBA project object model" in Excel (see Excel Setup Requirements above).

### Excel file created but modules not imported
The PowerShell script may have failed. Check that:
- Excel is installed and accessible
- You have permission to create files in C:\VBA_Projects
- Run the import manually: Open Excel, run `LocalUtility.ImportAllModules`

## Custom Templates

To create your own template:

1. Create a directory in `.vba-tools/templates/mytemplate/`
2. Add `template.json` with folder structure definition
3. Add `LocalUtility.bas` with placeholders: `{{WSL_WORKBOOK_PATH}}`, `{{WINDOWS_WORKBOOK_PATH}}`, `{{WORKBOOK_NAME}}`

Example template.json:
```json
{
  "name": "mytemplate",
  "description": "My custom template",
  "folders": [
    "Modules",
    "Classes",
    "Utils"
  ]
}
```

## Advanced Usage

### Create project with custom Windows path
```bash
vba-tool new MyProject --windows-path "D:\\Projects\\MyProject"
```

### Create multi-workbook project
```bash
vba-tool new SalesSystem --multi
# Then add workbooks:
vba-tool add-workbook SalesSystem Invoices
vba-tool add-workbook SalesSystem Reports
```

## Version Control

Projects are automatically git-initialized (unless --no-git is used).

The .gitignore includes:
- LocalUtility.bas files (development-specific)
- Excel temp files (~$*)
- Backup files (*.bak, *.tmp)

LocalUtility.bas files are local to each developer and contain machine-specific paths.
