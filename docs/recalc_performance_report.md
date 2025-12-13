# 再計算パフォーマンス計測レポート

## 前提
- ブック生成: `python tools/build_workbook.py --projects {10|30|60} --sample-all --output ModernExcelPMS_{N}.xlsx`
- 計測環境: Windows 11 / Microsoft 365 Excel 2308、メモリ 32GB。
- 操作手順:
  1. `Kanban_View!B2` で `PRJ_001` を選択し、`To Do / Doing / Done` の 3 カード表示を発火させる。
  2. `Measure_Master!F2` に `PRJ_001` を入力し、`INDIRECT` を通じた進捗参照を有効化する。
  3. `docs/vba/manual_calculation.bas` をインポートし、`MeasureFullRecalcWithLog` を実行して `Application.CalculateFullRebuild` の経過秒数を取得する。
  4. 手順 1–3 を 3 回繰り返し、中央値を採用する。

## 計測結果
| PRJ シート数 | 計算モード | `CalculateFullRebuild` 所要時間 (中央値) | 体感メモ |
| --- | --- | --- | --- |
| 10 | 自動 | 約 0.7 秒 | 画面の一瞬のフリーズのみで編集継続可 |
| 30 | 自動 | 約 2.1 秒 | `Kanban_View!B2` 編集後に待ちが発生、許容範囲 |
| 60 | 自動 | 約 4.3 秒 | `INDIRECT` のボラティリティが効き、数秒待つケースあり |
| 30 | 手動→フル再計算 | 約 2.0 秒 | 編集中はラグ無し。必要箇所で `MeasureFullRecalcWithLog` を呼ぶ運用で快適 |
| 60 | 手動→フル再計算 | 約 4.1 秒 | 編集中は軽快だが、再計算ボタン押下時にまとまって待つ形になる |

## 推奨
- 30 シートまでは自動計算で常用して問題なし。
- 31〜50 シートは編集負荷が高いフェーズだけ `ToggleManualCalculation` で手動に切り替え、節目で `MeasureFullRecalcWithLog` を実行する。
- 60 シート超ではブック分割または案件単位での管理を推奨。分割後は `Measure_Master` の参照先を最新ブックに限定する。

## 手動再計算の運用手順
1. Alt+F8 から `ToggleManualCalculation` を実行し、ステータスバーが「計算モード: 手動」となることを確認する。
2. 行のコピーやドラッグなど編集をまとめて実施する（自動再計算は発火しない）。
3. 適宜 `MeasureFullRecalcWithLog` を実行して所要時間を確認しつつ計算する。
4. 保存前に再度 `ToggleManualCalculation` を実行し、自動に戻した上で `MeasureFullRecalcWithLog` を 1 回流してから保存する。
