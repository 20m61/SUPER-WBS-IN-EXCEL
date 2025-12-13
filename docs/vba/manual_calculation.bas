Attribute VB_Name = "modManualCalculation"
Option Explicit

' 自動 / 手動計算をトグルし、ステータスバーへ反映する
Public Sub ToggleManualCalculation()
    With Application
        If .Calculation = xlCalculationAutomatic Then
            .Calculation = xlCalculationManual
            .StatusBar = "計算モード: 手動"
        Else
            .Calculation = xlCalculationAutomatic
            .CalculateFullRebuild
            .StatusBar = False
        End If
    End With
End Sub

' フル再計算の所要時間をイミディエイトウィンドウへ出力
Public Sub MeasureFullRecalcWithLog()
    Dim started As Double
    started = Timer
    Application.CalculateFullRebuild
    Debug.Print "再計算完了: " & Format(Timer - started, "0.000") & " 秒"
End Sub

' 複数シート数のブックで一括計測する場合のサンプルラッパー
Public Sub RunFullRecalcSuite()
    Dim i As Long
    For i = 1 To 3
        Debug.Print "--- 計測 " & i & " 回目 ---"
        MeasureFullRecalcWithLog
    Next i
End Sub
