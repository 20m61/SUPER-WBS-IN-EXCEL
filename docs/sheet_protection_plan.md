# シート保護設計とパスワード運用方針

Modern Excel PMS でユーザーが編集できる列と保護したい列を整理し、`.xlsm` 生成時に適用する保護設定とパスワード運用をまとめる。

## 1. 編集可 / 保護対象の整理
| シート | ユーザー編集列・操作 | 保護対象（ロックするセル/操作） | 備考 |
| --- | --- | --- | --- |
| Config | 祝日 `B4:B200`、担当者マスタ `D4:D200`、ステータス候補 `F4:F200` | セル見出し（1〜3 行目）、名前付き範囲の参照先調整用の空行以外の構造 | マスタ入力を残すため行挿入は許可。 |
| Template / PRJ_xxx | `Lv`/`タスク名`/`担当`/`開始日`/`工数(日)`/`進捗率`/`ステータス`/`備考` (A〜E,G,H,I 列)、タスク行の挿入・削除 | `終了日` (F 列)、全体進捗 `J2`、ヘッダー (4 行目)、ガント領域（F〜J 列以降の条件付き書式範囲）、計算ヘッダー | 行操作ボタン用のシェイプを保護するため DrawingObjects を保護対象に含める。 |
| Case_Master | 案件 ID/名前 (A,B 列)、メモ (C 列) | `施策数`/`平均進捗` (D,E 列)、ドリルダウン領域 (G2:N104) とプルダウン `H1` | フィルター・並べ替えは許可。 |
| Measure_Master | 施策 ID (A 列)、親案件 ID (B 列)、施策名 (C 列)、開始日 (D 列)、WBS シート名 (F 列)、備考 (H 列) | `WBS リンク` (E 列)、`実進捗` (G 列)、ヘッダー行 | B 列のデータ検証を維持するため保護状態でも入力可に設定。 |
| Kanban_View | WBS シート選択 `B2` | カード生成式 (B5:G104)、ヘッダー (1〜4 行) | ダブルクリックイベントでのみ更新させるため、セル編集は `B2` のみに限定。 |

### 推奨オプション
* 既定でセルをロックし、上記「編集可」セルのみロック解除してからシート保護を有効化する。
* 保護オプション: `AllowFormattingCells`, `AllowSorting`, `AllowFiltering`, `AllowInsertingRows`（Template/PRJ_xxx のみ）を有効にし、その他は無効。
* `UserInterfaceOnly:=True` を VBA で設定し、マクロ操作時の解除を不要にする。

## 2. `.xlsm` 生成時の保護適用手順（スクリプト設計）
1. **ロック状態の書き出し**: `tools/build_workbook.py` のセル定義に「ロック/解除」のメタ情報を追加し、`styles.xml` に `locked=0` のセル XF を定義して編集可能セルに付与する。
2. **シート保護ノードの追加**: 各 `worksheet` XML に `<sheetProtection>` を追記する。パスワードは Excel の XOR 方式（例: `hash = "DAA7"`）で事前計算して埋め込む。許可フラグは `formatCells="1" sort="1" autoFilter="1" insertRows="1"` などシート別に設定する。
3. **マクロ有効化**: `.xlsx` 出力後に VBA プロジェクトを追加し `.xlsm` として書き出す。手順例:
   1. 開発用の空 VBA プロジェクトを含むテンプレートブックから `vbaProject.bin` を抽出し、`zipfile` で `xl/vbaProject.bin` として同梱する。
   2. `[Content_Types].xml` に `application/vnd.ms-office.vbaProject` の Override を追加する。
4. **ビルド時のパスワード注入**: パスワードを環境変数 `PMS_SHEET_PASSWORD` で受け取り、未指定時は既定値を使用する。ハッシュ計算を Python で行い、`sheetProtection` に挿入する。
5. **検証**: 生成した `.xlsm` を Excel で開き、保護の有効性、行挿入可否、データ検証の動作を確認する回帰チェックを CI に組み込む（手動手順でも可）。

## 3. VBA での解除・再適用ユーティリティ
`.xlsm` 内に標準モジュール `modProtection` を追加し、次のユーティリティで保護を制御する。

```vba
Option Explicit
Private Const PROTECT_PASSWORD As String = "pms-2024"

Public Sub UnprotectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=PROTECT_PASSWORD
        On Error GoTo 0
    Next ws
End Sub

Public Sub ProtectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect _
            Password:=PROTECT_PASSWORD, _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            AllowFormattingCells:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowInsertingRows:=(ws.Name = "Template" Or ws.Name Like "PRJ_*")
        ws.EnableSelection = xlUnlockedCells
    Next ws
End Sub

Public Sub ReapplyProtection()
    UnprotectAllSheets
    ProtectAllSheets
End Sub
```

* `UserInterfaceOnly` は Excel が保存時に保持しないため、Workbook_Open イベントで `ProtectAllSheets` を呼び直すか、シート保護後に `UserInterfaceOnly:=True` を再設定する初期化マクロを追加する。 

## 4. 保護パスワードの運用
* **共有方法**: 1Password などの共有ボールトに「Modern Excel PMS シート保護パスワード」として登録し、閲覧権限を開発者と運用担当に限定する。メール・チャットでの平文共有は禁止。
* **変更手順**:
  1. 保護を解除した上で `PROTECT_PASSWORD` 定数と環境変数（CI 設定）を新しい値に更新。
  2. `build_workbook.py` のハッシュ生成に新パスワードを反映させて `.xlsm` を再生成。
  3. 旧パスワードをボールトから無効化し、周知をドキュメント（このファイル）と運用手順書に反映。
* **緊急時**: パスワード紛失時は VBA の `UnprotectAllSheets` をコメントアウトしつつブレークポイントで `PROTECT_PASSWORD` を再設定して復旧し、ローテーション後に再配布する。
