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


' Import all VBA modules from all project folders
Public Sub ImportAllModules()
    Dim wb As Workbook
    Dim vbProj As Object
    Dim fso As Object
    Dim allFiles As Collection
    Dim filePath As Variant
    Dim importedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    Dim startTime As Double

    On Error GoTo ErrorHandler

    startTime = Timer
    Set wb = ActiveWorkbook
    Set vbProj = wb.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set allFiles = New Collection

    Debug.Print "=== ImportAllModules Started ==="
    Debug.Print "Base Path: " & BASE_PATH

    ' Collect all .bas files from all folders
    Call CollectBasFiles(fso, BASE_PATH & "BaseSheets", allFiles)
    Call CollectBasFiles(fso, BASE_PATH & "WorkbookOperations", allFiles)
    Call CollectBasFiles(fso, BASE_PATH & "Utilities", allFiles)
    Call CollectBasFiles(fso, BASE_PATH & "Tests", allFiles)
    Call CollectBasFiles(fso, BASE_PATH & "Config", allFiles)

    Debug.Print "Found " & allFiles.Count & " .bas files to import"

    ' Import each file
    importedCount = 0
    skippedCount = 0
    errorCount = 0

    For Each filePath In allFiles
        If ModuleExists(vbProj, GetModuleName(CStr(filePath))) Then
            Debug.Print "  SKIPPED (already exists): " & GetModuleName(CStr(filePath))
            skippedCount = skippedCount + 1
        Else
            On Error Resume Next
            vbProj.VBComponents.Import CStr(filePath)
            If Err.Number = 0 Then
                Debug.Print "  IMPORTED: " & GetModuleName(CStr(filePath))
                importedCount = importedCount + 1
            Else
                Debug.Print "  ERROR: " & GetModuleName(CStr(filePath)) & " - " & Err.Description
                errorCount = errorCount + 1
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next filePath

    Debug.Print "=== ImportAllModules Complete ==="
    Debug.Print "Total Files: " & allFiles.Count
    Debug.Print "Imported: " & importedCount
    Debug.Print "Skipped (existing): " & skippedCount
    Debug.Print "Errors: " & errorCount
    Debug.Print "Time: " & Format(Timer - startTime, "0.00") & " seconds"

    Exit Sub

ErrorHandler:
    MsgBox "Error importing modules: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure 'Trust access to the VBA project object model' is enabled in:" & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings > Macro Settings", _
           vbCritical, "Import Error"
    Debug.Print "ERROR: " & Err.Description
End Sub


' Recursively collect all .bas files from a folder
Private Sub CollectBasFiles(fso As Object, folderPath As String, collection As Collection)
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object

    On Error Resume Next
    Set folder = fso.GetFolder(folderPath)

    If folder Is Nothing Then
        Debug.Print "WARNING: Folder not found: " & folderPath
        Exit Sub
    End If
    On Error GoTo 0

    ' Add all .bas files in this folder
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Path)) = "bas" Then
            collection.Add file.Path
        End If
    Next file

    ' Recursively process subfolders
    For Each subFolder In folder.SubFolders
        Call CollectBasFiles(fso, subFolder.Path, collection)
    Next subFolder
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


' Extract module name from file path
Private Function GetModuleName(filePath As String) As String
    Dim fso As Object
    Dim fileName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetFileName(filePath)

    ' Remove .bas extension
    GetModuleName = Left(fileName, Len(fileName) - 4)
End Function


' ======================================================================
' RAPID DEVELOPMENT UTILITIES
' ======================================================================


' Remove all modules and re-import them from disk
Public Sub ReloadAllModules()
    Const vbext_ct_StdModule As Long = 1
    Dim wb As Workbook
    Dim vbProj As Object
    Dim comp As Object
    Dim moduleName As String
    Dim removedCount As Long
    Dim startTime As Double
    Dim i As Long

    On Error GoTo ErrorHandler

    startTime = Timer
    Set wb = ActiveWorkbook
    Set vbProj = wb.VBProject

    Debug.Print "=== ReloadAllModules Started ==="
    Debug.Print "Removing all existing modules..."
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
    Debug.Print "Re-importing all modules..."
    Call ImportAllModules

    Debug.Print "=== ReloadAllModules Complete ==="
    Debug.Print "Total time: " & Format(Timer - startTime, "0.00") & " seconds"

    Exit Sub

ErrorHandler:
    MsgBox "Error reloading modules: " & Err.Description, vbCritical, "Reload Error"
    Debug.Print "ERROR in ReloadAllModules: " & Err.Description
End Sub


' Reload only development modules (Directives/Overview) for fast iteration
Public Sub ReloadDevModules()
    Const vbext_ct_StdModule As Long = 1
    Dim wb As Workbook
    Dim vbProj As Object
    Dim comp As Object
    Dim fso As Object
    Dim allFiles As Collection
    Dim filePath As Variant
    Dim moduleName As String
    Dim removedCount As Long
    Dim importedCount As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    Set wb = ActiveWorkbook
    Set vbProj = wb.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set allFiles = New Collection

    Debug.Print "=== ReloadDevModules Started ==="
    removedCount = 0

    ' Remove Directive and Overview modules
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents(i)

        If comp.Type = vbext_ct_StdModule Then
            moduleName = comp.Name

            If Left(moduleName, 10) = "Directive_" Or _
               Left(moduleName, 9) = "Overview_" Then
                On Error Resume Next
                vbProj.VBComponents.Remove comp
                If Err.Number = 0 Then
                    Debug.Print "  REMOVED: " & moduleName
                    removedCount = removedCount + 1
                End If
                On Error GoTo ErrorHandler
            End If
        End If
    Next i

    Debug.Print "Removed " & removedCount & " modules"

    ' Re-import from Directives and Overview folders
    Call CollectBasFiles(fso, BASE_PATH & "WorkbookOperations\Directives", allFiles)
    Call CollectBasFiles(fso, BASE_PATH & "WorkbookOperations\Overview", allFiles)

    importedCount = 0
    For Each filePath In allFiles
        If Not ModuleExists(vbProj, GetModuleName(CStr(filePath))) Then
            On Error Resume Next
            vbProj.VBComponents.Import CStr(filePath)
            If Err.Number = 0 Then
                Debug.Print "  IMPORTED: " & GetModuleName(CStr(filePath))
                importedCount = importedCount + 1
            End If
            On Error GoTo ErrorHandler
        End If
    Next filePath

    Debug.Print "=== ReloadDevModules Complete ==="
    Debug.Print "Imported: " & importedCount

    Exit Sub

ErrorHandler:
    MsgBox "Error reloading dev modules: " & Err.Description, vbCritical, "Reload Error"
    Debug.Print "ERROR in ReloadDevModules: " & Err.Description
End Sub


' Run all test modules
Public Sub RunAllTests()
    Dim vbProj As Object
    Dim comp As Object
    Dim moduleName As String
    Dim testsRun As Long
    Dim testsFailed As Long

    On Error GoTo ErrorHandler

    Set vbProj = ActiveWorkbook.VBProject

    Debug.Print "=== Running All Tests ==="
    testsRun = 0
    testsFailed = 0

    ' Find and run all Test_ modules
    For Each comp In vbProj.VBComponents
        If comp.Type = 1 Then ' Standard module
            moduleName = comp.Name

            If Left(moduleName, 5) = "Test_" Then
                Debug.Print "Running: " & moduleName

                On Error Resume Next
                Application.Run moduleName & ".RunTests"

                If Err.Number <> 0 Then
                    Debug.Print "  FAILED: " & Err.Description
                    testsFailed = testsFailed + 1
                    Err.Clear
                Else
                    Debug.Print "  PASSED"
                End If

                On Error GoTo ErrorHandler
                testsRun = testsRun + 1
            End If
        End If
    Next comp

    Debug.Print "=== Tests Complete ==="
    Debug.Print "Tests Run: " & testsRun
    Debug.Print "Passed: " & (testsRun - testsFailed)
    Debug.Print "Failed: " & testsFailed

    Exit Sub

ErrorHandler:
    Debug.Print "ERROR running tests: " & Err.Description
End Sub
