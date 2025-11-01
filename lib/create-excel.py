#!/usr/bin/env python3

"""
Excel File Creator for VBA Tool
Creates .xlsm files on Windows and imports LocalUtility.bas module
"""

import sys
import os
import subprocess
import json
from pathlib import Path

def create_powershell_script(excel_path, vba_module_path):
    """Generate PowerShell script to create Excel file and import VBA"""

    # Convert WSL paths to Windows paths for PowerShell
    if vba_module_path.startswith('/mnt/'):
        # /mnt/c/... -> C:\...
        vba_module_path_win = vba_module_path.replace('/mnt/c/', 'C:\\').replace('/', '\\')
    elif vba_module_path.startswith('/home/'):
        # /home/... -> \\wsl.localhost\Ubuntu\home\...
        vba_module_path_win = f"\\\\wsl.localhost\\Ubuntu{vba_module_path}"
    else:
        vba_module_path_win = vba_module_path

    ps_script = f"""
# Excel File Creation Script
$ErrorActionPreference = "Stop"

try {{
    Write-Host "Creating Excel file: {excel_path}"

    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Create new workbook
    $workbook = $excel.Workbooks.Add()

    # Enable VBA project access (required for importing modules)
    # Note: User must have "Trust access to VBA project object model" enabled in Excel settings

    # Import LocalUtility.bas if it exists
    $vbaModulePath = "{vba_module_path_win}"

    if (Test-Path $vbaModulePath) {{
        Write-Host "Importing VBA module: $vbaModulePath"
        try {{
            $workbook.VBProject.VBComponents.Import($vbaModulePath)
            Write-Host "VBA module imported successfully"
        }} catch {{
            Write-Host "WARNING: Could not import VBA module. Make sure 'Trust access to VBA project object model' is enabled."
            Write-Host "Error: $_"
        }}
    }} else {{
        Write-Host "WARNING: VBA module not found at: $vbaModulePath"
    }}

    # Save as macro-enabled workbook
    $workbook.SaveAs("{excel_path}", 52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)

    Write-Host "Excel file created successfully: {excel_path}"

    # Close workbook and Excel
    $workbook.Close($false)
    $excel.Quit()

    # Clean up COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    exit 0
}} catch {{
    Write-Host "ERROR: $_"
    if ($excel) {{
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }}
    exit 1
}}
"""
    return ps_script


def main():
    if len(sys.argv) < 3:
        print("Usage: create-excel.py <excel_path> <vba_module_path>")
        print("Example: create-excel.py 'C:\\VBA_Projects\\MyProject\\MyWorkbook.xlsm' '/home/user/vba/MyProject/LocalUtility.bas'")
        sys.exit(1)

    excel_path = sys.argv[1]
    vba_module_path = sys.argv[2]

    # Validate paths
    if not os.path.exists(vba_module_path):
        print(f"ERROR: VBA module not found: {vba_module_path}")
        sys.exit(1)

    # Create directory for Excel file if it doesn't exist
    excel_dir = os.path.dirname(excel_path)
    if '\\' in excel_dir:
        # Convert to WSL path for directory creation
        excel_dir_wsl = excel_dir.replace('C:\\', '/mnt/c/').replace('\\', '/')
        os.makedirs(excel_dir_wsl, exist_ok=True)

    # Generate PowerShell script
    ps_script = create_powershell_script(excel_path, vba_module_path)

    # Write PowerShell script to temp file
    temp_ps_file = "/tmp/create-excel.ps1"
    with open(temp_ps_file, 'w', newline='\r\n') as f:
        f.write(ps_script)

    # Execute PowerShell script via powershell.exe
    try:
        # Convert temp file to Windows path
        temp_ps_file_win = subprocess.check_output(
            ['wslpath', '-w', temp_ps_file],
            text=True
        ).strip()

        print(f"Executing PowerShell script to create Excel file...")

        # Run PowerShell script
        result = subprocess.run(
            ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', temp_ps_file_win],
            capture_output=True,
            text=True
        )

        # Print output
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr, file=sys.stderr)

        if result.returncode != 0:
            print(f"ERROR: PowerShell script failed with return code {result.returncode}")
            sys.exit(1)

        print("Excel file created successfully!")

    except subprocess.CalledProcessError as e:
        print(f"ERROR: Failed to execute PowerShell script: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
    finally:
        # Clean up temp file
        if os.path.exists(temp_ps_file):
            os.remove(temp_ps_file)


if __name__ == "__main__":
    main()
