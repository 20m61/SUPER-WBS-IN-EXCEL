Attribute VB_Name = "modProtection"
Option Explicit
Private Const PROTECT_PASSWORD As String = "pms-2024"

' 全シートの保護を解除する
Public Sub UnprotectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=PROTECT_PASSWORD
        On Error GoTo 0
    Next ws
End Sub

' 編集可能セルだけを解放した状態で保護を再適用する
' UserInterfaceOnly:=True によりマクロからの操作は保護を無視できる
Public Sub ProtectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' 一度解除してから再適用（UserInterfaceOnly を確実に設定するため）
        On Error Resume Next
        ws.Unprotect Password:=PROTECT_PASSWORD
        On Error GoTo 0

        ws.Protect _
            Password:=PROTECT_PASSWORD, _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            UserInterfaceOnly:=True, _
            AllowFormattingCells:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowInsertingRows:=(ws.Name = "Template" Or ws.Name Like "PRJ_*")
        ws.EnableSelection = xlUnlockedCells
    Next ws
End Sub

' 解除→再適用のラッパー。初期化マクロや設定変更時に利用する
Public Sub ReapplyProtection()
    UnprotectAllSheets
    ProtectAllSheets
End Sub
