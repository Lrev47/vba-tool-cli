Attribute VB_Name = "ProjectEntry"
Option Explicit

' Main entry point for {{WORKBOOK_NAME}}
' This macro is called by the "Run {{WORKBOOK_NAME}}" button on Sheet1
'
' This module is version-controlled and gets reloaded by ResetWorkbook()
' Edit this file to add your project logic

Public Sub RunProject()
    ' Add your project logic here
    ' Call other modules, orchestrate your workflow

    MsgBox "{{WORKBOOK_NAME}} is running!" & vbCrLf & vbCrLf & _
           "Edit ProjectEntry.bas to implement your logic." & vbCrLf & vbCrLf & _
           "The button is wired to ProjectEntry.RunProject", _
           vbInformation, "{{WORKBOOK_NAME}}"
End Sub
