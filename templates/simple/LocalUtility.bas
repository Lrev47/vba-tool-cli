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
