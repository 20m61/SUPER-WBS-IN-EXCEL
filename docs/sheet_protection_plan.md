# シート保護設計とパスワード運用方針

Modern Excel PMS でユーザーが編集できる列と保護したい列を整理し、`.xlsm` 生成時に適用する保護設定とパスワード運用をまとめる。

## 1. 編集可 / 保護対象の整理
| シート | ユーザー編集列・操作 | 保護対象（ロックするセル/操作） | 備考 |
| --- | --- | --- | --- |
| Config | 祝日 `B4:B200`、担当者マスタ `D4:D200`、ステータス候補 `F4:F200` | 見出し行（1〜3 行目）、命名済み範囲の参照先、空行以外の表構造 | マスタ入力を残すため行挿入は許可。 |
| Template / PRJ_xxx | `Lv`/`タスク名`/`担当`/`開始日`/`工数(日)`/`進捗率`/`ステータス`/`備考` (A〜E,G,H,I 列)、タスク行の挿入・削除 | `終了日` (F 列)、全体進捗 `J2`、ヘッダー (4 行目)、ガント領域（F〜J 列以降の条件付き書式範囲）、計算ヘッダー | 行操作ボタン用のシェイプを保護するため DrawingObjects を保護対象に含める。 |
| Case_Master | 案件 ID/名前 (A,B 列)、メモ (C 列) | `施策数`/`平均進捗` (D,E 列)、ドリルダウン領域 (G2:N104) とプルダウン `H1` | フィルター・並べ替えは許可。 |
| Measure_Master | 施策 ID (A 列)、親案件 ID (B 列)、施策名 (C 列)、開始日 (D 列)、WBS シート名 (F 列)、備考 (H 列) | `WBS リンク` (E 列)、`実進捗` (G 列)、ヘッダー行 | B 列のデータ検証を維持するため保護状態でも入力可に設定。 |
| Kanban_View | WBS シート選択 `B2` | カード生成式 (B5:G104)、ヘッダー (1〜4 行) | ダブルクリックイベントでのみ更新させるため、セル編集は `B2` のみに限定。 |

### ロック/ロック解除の考え方（再掲）
* 既定で全セルをロックし、上表の「ユーザー編集列」に対してのみロック解除を設定する。
* 保護オプション: `AllowFormattingCells`, `AllowSorting`, `AllowFiltering` を既定で有効化し、Template/PRJ_xxx のみ `AllowInsertingRows` を追加。行削除は VBA で制御するため許可しない。
* `UserInterfaceOnly:=True` を VBA で設定し、マクロ操作時の解除を不要にする。保存時に失われるため Workbook_Open で再設定する。

### 推奨オプション
* 既定でセルをロックし、上記「編集可」セルのみロック解除してからシート保護を有効化する。
* 保護オプション: `AllowFormattingCells`, `AllowSorting`, `AllowFiltering`, `AllowInsertingRows`（Template/PRJ_xxx のみ）を有効にし、その他は無効。
* `UserInterfaceOnly:=True` を VBA で設定し、マクロ操作時の解除を不要にする。

## 2. `.xlsm` 生成時の保護適用手順（スクリプト設計）
1. **ロック状態のメタ情報定義**  
   `tools/build_workbook.py` のセル定義に `locked: bool` を保持するタプルを追加する想定で、書式スタイルに `locked=0` を持つ XF (例: styleId=1) を定義する。編集可能セルはこの XF を参照し、その他はデフォルト XF (styleId=0) を適用する。
2. **シート別に保護フラグを付与**  
   `worksheet_xml` に `<sheetProtection>` を埋め込み、以下のフラグをシートごとに設定する。
   * 共通: `sheet="1"`、`objects="1"`（DrawingObjects を含む）、`formatCells="1"`、`sort="1"`、`autoFilter="1"`
   * Template/PRJ_xxx のみ: 上記に加え `insertRows="1"`
   パスワードは Excel の XOR ハッシュを Python で計算し、`password="XXXX"` に埋め込む。
3. **VBA プロジェクトの同梱**  
   `.xlsx` を書き出した後に `xl/vbaProject.bin` を追加し、`[Content_Types].xml` に `application/vnd.ms-office.vbaProject` の Override を追加して `.xlsm` として完成させる。`zipfile` の再圧縮時に既存パーツを壊さないよう一時ディレクトリで展開→再パックするステップをスクリプトに組み込む。
4. **パスワードの受け渡しと再生成**  
   環境変数 `PMS_SHEET_PASSWORD` を優先し、未設定時はリポジトリ既定値（例: `pms-2024`）を使う。変更時は `build_workbook.py` が同じ関数でハッシュを再計算することで `.xlsm` を自動再生成できるようにする。
5. **検証フロー**  
   CI もしくは手動で、(a) PRJ シートの行挿入可否、(b) Measure_Master のデータ検証維持、(c) ガント列など保護列の編集不可、(d) VBA からの編集可否（UserInterfaceOnly 動作）を確認するテストシナリオをリスト化し、ビルド成果物ごとに実施する。

## 3. VBA での解除・再適用ユーティリティ
`.xlsm` に標準モジュール `modProtection` を同梱し、シート保護の一括操作を提供する。`UserInterfaceOnly` を確実に適用するため Workbook_Open から `ProtectAllSheets` を呼び出す運用を前提とする。

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

### VBA モジュールの配置案
* `modProtection.bas`: 上記ユーティリティ本体。標準モジュールとしてプロジェクトに含める。
* `ThisWorkbook`: `Workbook_Open` イベントで `ProtectAllSheets` を呼び出し、保存後も保護状態を復元する。
* 将来的に行操作やテンプレート複製マクロ（`modWbsCommands`）を追加する際も、保護解除を必要としないよう `UserInterfaceOnly` を維持したまま操作する。

## 4. 保護パスワードの運用
* **共有方法**: 1Password などの共有ボールトに「Modern Excel PMS シート保護パスワード」として登録し、閲覧権限を開発・運用担当に限定する。メール・チャットでの平文共有は禁止。
* **変更手順（開発〜配布までの統一フロー）**:
  1. 既存ブックで保護を解除してから、新パスワードをボールトに登録。
  2. `modProtection.bas` の `PROTECT_PASSWORD` 定数と CI 環境変数 `PMS_SHEET_PASSWORD` を更新し、`build_workbook.py` のハッシュ計算に新パスワードが反映されることを確認。
  3. `.xlsm` を再生成し、サンプルブックで保護設定が新パスワードで動作するか検証。
  4. 旧パスワードのエントリをボールトで無効化し、変更履歴をこのドキュメントと運用手順書に記載。
* **緊急時**: パスワード紛失時は `UnprotectAllSheets` を一時的に無効化した上でデバッグウィンドウから `PROTECT_PASSWORD` を再設定し、直後にパスワードをローテーションして再配布する。ボールト外での共有は禁止。
