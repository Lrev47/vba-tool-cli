Attribute VB_Name = "LocalUtility"
Option Explicit

' Local development utility module for {{WORKBOOK_NAME}}
' This file is ignored by git - use for personal testing/development

' Requires: Microsoft Visual Basic for Applications Extensibility 5.3
' Requires: Trust access to the VBA project object model (enabled in Excel Trust Center)

' Auto-configured paths
' WSL Path: {{WSL_WORKBOOK_PATH}}
' Windows Path: {{WINDOWS_WORKBOOK_PATH}}

Const BASE_PATH As String = "{{WINDOWS_WORKBOOK_PATH}}\"


' Import all VBA modules from the workbook directory
Public Sub ImportAllModules()
    Dim wb As Workbook
    Dim vbProj As Object
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim importedCount As Long
    Dim skippedCount As Long
    Dim startTime As Double

    On Error GoTo ErrorHandler

    startTime = Timer
    Set wb = ActiveWorkbook
    Set vbProj = wb.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")

    Debug.Print "=== ImportAllModules Started ==="
    Debug.Print "Base Path: " & BASE_PATH

    If Not fso.FolderExists(BASE_PATH) Then
        MsgBox "Base path not found: " & BASE_PATH, vbCritical, "Path Error"
        Exit Sub
    End If

    Set folder = fso.GetFolder(BASE_PATH)

    importedCount = 0
    skippedCount = 0

    ' Import all .bas files in the directory
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Path)) = "bas" Then
            Dim moduleName As String
            moduleName = fso.GetBaseName(file.Name)

            ' Skip LocalUtility (this module)
            If moduleName <> "LocalUtility" Then
                If ModuleExists(vbProj, moduleName) Then
                    Debug.Print "  SKIPPED (already exists): " & moduleName
                    skippedCount = skippedCount + 1
                Else
                    On Error Resume Next
                    vbProj.VBComponents.Import file.Path
                    If Err.Number = 0 Then
                        Debug.Print "  IMPORTED: " & moduleName
                        importedCount = importedCount + 1
                    Else
                        Debug.Print "  ERROR: " & moduleName & " - " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                End If
            End If
        End If
    Next file

    Debug.Print "=== ImportAllModules Complete ==="
    Debug.Print "Imported: " & importedCount
    Debug.Print "Skipped: " & skippedCount
    Debug.Print "Time: " & Format(Timer - startTime, "0.00") & " seconds"

    Exit Sub

ErrorHandler:
    MsgBox "Error importing modules: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure 'Trust access to the VBA project object model' is enabled in:" & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings > Macro Settings", _
           vbCritical, "Import Error"
    Debug.Print "ERROR: " & Err.Description
End Sub


' Check if a module with given name already exists in the project
Private Function ModuleExists(vbProj As Object, moduleName As String) As Boolean
    Dim comp As Object

    ModuleExists = False

    On Error Resume Next
    Set comp = vbProj.VBComponents(moduleName)
    ModuleExists = Not comp Is Nothing
    On Error GoTo 0
End Function


' Remove all modules (except LocalUtility) and re-import them
Public Sub ReloadAllModules()
    Const vbext_ct_StdModule As Long = 1
    Dim wb As Workbook
    Dim vbProj As Object
    Dim comp As Object
    Dim moduleName As String
    Dim removedCount As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    Set wb = ActiveWorkbook
    Set vbProj = wb.VBProject

    Debug.Print "=== ReloadAllModules Started ==="

    ' Remove all standard modules except LocalUtility
    removedCount = 0
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents(i)

        If comp.Type = vbext_ct_StdModule Then
            moduleName = comp.Name

            If moduleName <> "LocalUtility" Then
                On Error Resume Next
                vbProj.VBComponents.Remove comp
                If Err.Number = 0 Then
                    Debug.Print "  REMOVED: " & moduleName
                    removedCount = removedCount + 1
                Else
                    Debug.Print "  ERROR removing " & moduleName & ": " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End If
        End If
    Next i

    Debug.Print "Removed " & removedCount & " modules"

    ' Re-import all modules
    Call ImportAllModules

    Debug.Print "=== ReloadAllModules Complete ==="

    Exit Sub

ErrorHandler:
    MsgBox "Error reloading modules: " & Err.Description, vbCritical, "Reload Error"
    Debug.Print "ERROR in ReloadAllModules: " & Err.Description
End Sub


' Reset workbook to clean state with Sheet1 and Run button
' Adapted from UsageWorkbook clear-rebuild pattern
Public Sub ResetWorkbook()
    Dim ws As Worksheet
    Dim sheet1 As Worksheet
    Dim btn As Button
    Dim projectName As String
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim i As Long

    On Error GoTo ErrorHandler

    ' Extract project name from workbook name
    projectName = Replace(ThisWorkbook.Name, ".xlsm", "")
    projectName = Replace(projectName, ".xls", "")

    ' Save application state and disable for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Debug.Print "=== ResetWorkbook Started ==="
    Debug.Print "Project: " & projectName

    ' Step 1: Reload all modules to pick up code changes
    Debug.Print "Step 1: Reloading modules..."
    Call ReloadAllModules

    ' Step 2: Make all sheets visible
    Debug.Print "Step 2: Making all sheets visible..."
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            On Error Resume Next
            ws.Visible = xlSheetVisible
            If Err.Number = 0 Then
                Debug.Print "  Made visible: " & ws.Name
            Else
                Debug.Print "  WARNING: Could not make visible: " & ws.Name
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next ws

    ' Step 3: Store reference to first sheet
    Set sheet1 = ThisWorkbook.Worksheets(1)
    Debug.Print "First sheet: " & sheet1.Name

    ' Step 4: Delete all sheets except first
    Debug.Print "Step 3: Deleting all sheets except first..."
    For i = ThisWorkbook.Worksheets.Count To 2 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        Debug.Print "  Deleting: " & ws.Name
        On Error Resume Next
        ws.Delete
        If Err.Number <> 0 Then
            Debug.Print "  ERROR deleting: " & ws.Name & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i

    ' Step 5: Ensure first sheet is named "Sheet1"
    Debug.Print "Step 4: Ensuring first sheet is named Sheet1..."
    On Error Resume Next
    If sheet1.Name <> "Sheet1" Then
        sheet1.Name = "Sheet1"
        If Err.Number = 0 Then
            Debug.Print "  Renamed to: Sheet1"
        Else
            Debug.Print "  WARNING: Could not rename - " & Err.Description
            Err.Clear
        End If
    Else
        Debug.Print "  Already named: Sheet1"
    End If
    On Error GoTo ErrorHandler

    ' Step 6: Clear all content and formatting (like UsageWorkbook EnsureAndClear)
    Debug.Print "Step 5: Clearing all content and formatting..."
    sheet1.Cells.Clear
    sheet1.Cells.ClearFormats
    Debug.Print "  Sheet cleared"

    ' Step 7: Recreate Run button in cell C1 (like UsageWorkbook pattern)
    Debug.Print "Step 6: Creating Run button in cell C1..."
    On Error Resume Next

    ' Position button to fill cell C1
    With sheet1.Range("C1")
        btnLeft = .Left
        btnTop = .Top
        btnWidth = .Width
        btnHeight = .Height
    End With

    ' Create button
    Set btn = sheet1.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)

    If Err.Number = 0 Then
        btn.Caption = "Run " & projectName
        btn.OnAction = "ProjectEntry.RunProject"  ' Wired to ProjectEntry module

        Debug.Print "  Button created successfully"
        Debug.Print "    Caption: " & btn.Caption
        Debug.Print "    Position: Cell C1"
        Debug.Print "    OnAction: ProjectEntry.RunProject"
    Else
        Debug.Print "  ERROR creating button: " & Err.Description
        Err.Clear
    End If

    On Error GoTo ErrorHandler

    ' Restore application state
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Activate Sheet1
    sheet1.Activate

    Debug.Print "=== ResetWorkbook Complete ==="

    MsgBox "Workbook reset successfully!" & vbCrLf & vbCrLf & _
           "Sheet1 is ready with '" & btn.Caption & "' button" & vbCrLf & vbCrLf & _
           "Button wired to: ProjectEntry.RunProject" & vbCrLf & _
           "Edit ProjectEntry.bas to add your logic", _
           vbInformation, "Reset Complete"

    Exit Sub

ErrorHandler:
    ' Always restore application state on error
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Debug.Print "ERROR in ResetWorkbook: " & Err.Description
    MsgBox "Error resetting workbook: " & Err.Description & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for details.", _
           vbCritical, "Reset Error"
End Sub
