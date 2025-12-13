## 再計算パフォーマンス検証メモ

### 目的
`INDIRECT` と `FILTER` を多数の PRJ シートに対して呼び出した際の再計算時間を把握し、自動計算で運用できる上限と、手動計算へ切り替える際の手順をまとめる。

### ブック生成手順
`tools/build_workbook.py` で任意枚数の PRJ シートを含む検証用ブックを作れる。`--sample-first` を付けると最初の PRJ シートにサンプルタスクを埋め込む。複数シート全てにダミータスクを置いて `FILTER` のヒット件数を安定させたい場合は `--sample-all` を利用する。

```bash
# 例: 30 枚の PRJ シートを持つブックを生成（全シートにサンプル行を配置）
python tools/build_workbook.py --projects 30 --sample-all --output ./ModernExcelPMS_30.xlsx
```

生成時間とファイルサイズはシート枚数にほぼ比例し、生成だけなら 60 枚でも数十 ms 程度で終わる。

| PRJ シート数 | 生成時間 | ファイルサイズ |
| --- | --- | --- |
| 10 | 約 0.008s | 約 0.02 MB |
| 30 | 約 0.011s | 約 0.04 MB |
| 60 | 約 0.019s | 約 0.07 MB |

### 再計算測定の進め方
1. 上記で生成したブックを Excel で開く。
2. `Kanban_View!B2` で複数の PRJ シートを順に選択し、`INDIRECT`/`FILTER` が多く発火する状態を再現する。
3. 直後に `Application.CalculateFullRebuild` を実行し、処理時間をストップウォッチまたは VBA で記録する。
   - `docs/vba/manual_calculation.bas` を標準モジュールとして取り込むと、計算モードの切り替えと経過時間記録が Alt+F8 から実行できる。
4. シート枚数を 10 → 30 → 60 と増やし、計算時間の伸びを比較する（ボラティリティのため線形増加が目安）。計測結果は `docs/recalc_performance_report.md` に追記する。

### 手動再計算トグル用 VBA の試作例
再計算を手動に切り替える場合は、以下のような標準モジュールを追加して `Alt+F8` から実行する。必要に応じてクイックアクセスツールバーへ登録すると便利。`docs/vba/manual_calculation.bas` に同等コードを用意しているのでインポートすれば入力不要。

```vba
Option Explicit

Public Sub ToggleManualCalculation()
    ' 自動→手動を切り替え、状態をステータスバーに表示する
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

Public Sub MeasureFullRecalc()
    ' 現在の計算モードで再計算し、経過時間をイミディエイトウィンドウに出力する
    Dim started As Double
    started = Timer
    Application.CalculateFullRebuild
    Debug.Print "再計算完了: " & Format(Timer - started, "0.000") & " 秒"
End Sub
```

### 推奨運用と最大シート数の目安
- `INDIRECT` はボラティルなため、PRJ シートが増えるほど `Kanban_View` と `Measure_Master` の再計算時間が直線的に伸びる。
- 30 枚程度までは自動計算でも体感待ち時間は短い想定。40 枚以上では計算トリガーを絞るため、上記マクロで手動計算へ切り替える運用を推奨。
- 60 枚を超えると `FILTER` の探索対象も増え、計算完了まで数秒かかる可能性が高い。長時間運用する場合は案件単位でブックを分割し、`Measure_Master` の参照を最新ブックのみに絞ると安定する。
- 手動計算時は、編集フェーズでは手動のまま作業し、要所で `CalculateFullRebuild` を実行する、保存前に自動へ戻してから再計算する、という手順が安全。
