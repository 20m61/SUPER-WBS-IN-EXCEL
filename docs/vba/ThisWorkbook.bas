Attribute VB_Name = "ThisWorkbook"
Option Explicit

' ブックを開いた際にシート保護を再適用する
' UserInterfaceOnly は保存時に失われるため、Open イベントで再設定が必要
Private Sub Workbook_Open()
    modProtection.ProtectAllSheets
End Sub

' 既存の PRJ_xxx シートを走査して次の採番を返す
Public Function NextProjectSheetName() As String
    Dim ws As Worksheet
    Dim maxIndex As Long
    Dim currentIndex As Long
    Dim baseName As String

    baseName = "PRJ_"
    maxIndex = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like baseName & "###" Then
            currentIndex = CLng(Mid$(ws.Name, 5))
            If currentIndex > maxIndex Then
                maxIndex = currentIndex
            End If
        End If
    Next ws

    NextProjectSheetName = baseName & Format$(maxIndex + 1, "000")
End Function
