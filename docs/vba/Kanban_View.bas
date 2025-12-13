Attribute VB_Name = "Kanban_View"
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ExitHandler

    If Intersect(Target, Range("B5:B104,D5:D104,F5:F104")) Is Nothing Then
        Exit Sub
    End If

    Cancel = True
    modWbsCommands.UpdateTaskStatusFromKanban Target
ExitHandler:
End Sub
