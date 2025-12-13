Attribute VB_Name = "modWbsCommands"
Option Explicit

Private Const TABLE_START_ROW As Long = 5
Private Const TABLE_END_ROW As Long = 104
Private Const TABLE_LAST_COLUMN As String = "I"

' 選択行を上に移動する
Public Sub MoveTaskRowUp()
    SwapTaskRow -1
End Sub

' 選択行を下に移動する
Public Sub MoveTaskRowDown()
    SwapTaskRow 1
End Sub

' Template から複製して次の PRJ_xxx を生成する
Public Sub DuplicateTemplateSheet()
    Dim nextName As String
    Dim templateSheet As Worksheet

    Set templateSheet = ThisWorkbook.Worksheets("Template")
    nextName = ThisWorkbook.NextProjectSheetName

    templateSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = nextName
End Sub

' カンバンから指定したタスクのステータスを更新する
Public Sub UpdateTaskStatusFromKanban(ByVal targetCell As Range)
    Dim sourceSheetName As String
    Dim targetStatus As String
    Dim taskName As String
    Dim taskSheet As Worksheet
    Dim found As Range
    Dim lines() As String

    sourceSheetName = targetCell.Parent.Range("B2").Value
    If sourceSheetName = "" Then
        MsgBox "WBS シート名が未選択のため更新できません。", vbExclamation
        Exit Sub
    End If

    Select Case targetCell.Column
        Case 2
            targetStatus = "未着手"
        Case 4
            targetStatus = "進行中"
        Case 6
            targetStatus = "完了"
        Case Else
            Exit Sub
    End Select

    If targetCell.Value = "" Then Exit Sub

    lines = Split(targetCell.Value, vbLf)
    taskName = Trim$(lines(0))
    If taskName = "" Then Exit Sub

    On Error Resume Next
    Set taskSheet = ThisWorkbook.Worksheets(sourceSheetName)
    On Error GoTo 0
    If taskSheet Is Nothing Then
        MsgBox "対象の WBS シートが見つかりません。", vbCritical
        Exit Sub
    End If

    Set found = taskSheet.Range("B" & TABLE_START_ROW & ":B" & TABLE_END_ROW).Find( _
        What:=taskName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If found Is Nothing Then
        MsgBox "タスク名が WBS 上に存在しません。", vbInformation
        Exit Sub
    End If

    taskSheet.Range("H" & found.Row).Value = targetStatus
End Sub

' 行のスワップを内部処理としてまとめる
Private Sub SwapTaskRow(ByVal direction As Long)
    Dim currentRow As Long
    Dim targetRow As Long
    Dim targetRange As Range
    Dim buffer As Variant
    Dim swapRange As String

    If ActiveCell Is Nothing Then Exit Sub
    currentRow = ActiveCell.Row

    If currentRow < TABLE_START_ROW Or currentRow > TABLE_END_ROW Then
        MsgBox "タスク表の行を選択してから実行してください。", vbInformation
        Exit Sub
    End If

    targetRow = currentRow + direction
    If targetRow < TABLE_START_ROW Or targetRow > TABLE_END_ROW Then
        MsgBox "これ以上移動できません。", vbInformation
        Exit Sub
    End If

    swapRange = "A" & currentRow & ":" & TABLE_LAST_COLUMN & currentRow
    buffer = ActiveSheet.Range(swapRange).Value

    swapRange = "A" & targetRow & ":" & TABLE_LAST_COLUMN & targetRow
    Application.ScreenUpdating = False
    Set targetRange = ActiveSheet.Range(swapRange)

    ActiveSheet.Range("A" & currentRow & ":" & TABLE_LAST_COLUMN & currentRow).Value = targetRange.Value
    targetRange.Value = buffer
    Application.ScreenUpdating = True

    ActiveSheet.Range("B" & targetRow).Select
End Sub
