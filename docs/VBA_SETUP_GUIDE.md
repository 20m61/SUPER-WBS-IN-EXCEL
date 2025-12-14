# VBA マクロセットアップガイド

VBAマクロ付きExcelファイルの生成と設定について説明します。

## 概要

`build_workbook.py` は `--with-vba` オプションを使用することで、VBAマクロを自動的に組み込んだ `.xlsm` ファイルを生成できます。

### 自動生成（推奨）

VBAプロジェクトバイナリ（vbaProject.bin）は自動生成されます。初回生成時にキャッシュされ、次回以降は高速に読み込まれます。

```bash
python3 tools/build_workbook.py \
  --output output/ModernExcelPMS.xlsm \
  --with-vba \
  --with-buttons \
  --sample-first
```

### VBAを再生成する場合

VBAソースコードを変更した場合は、`--regenerate-vba` オプションを使用してキャッシュを更新します。

```bash
python3 tools/build_workbook.py \
  --output output/ModernExcelPMS.xlsm \
  --with-vba \
  --with-buttons \
  --regenerate-vba \
  --sample-first
```

## 手動セットアップ（オプション）

自動生成がうまく動作しない場合は、以下の手順でVBAを手動追加できます。

### 前提条件

- Microsoft Excel（マクロ有効）
- 生成済みの `.xlsm` ファイル（`--with-buttons` オプション使用）

### 1. ファイルの生成（VBAなし）

```bash
python3 tools/build_workbook.py \
  --output output/ModernExcelPMS.xlsm \
  --with-buttons \
  --sample-first
```

### 2. Excelでファイルを開く

1. 生成された `.xlsm` ファイルをExcelで開く
2. セキュリティ警告が表示された場合は「コンテンツの有効化」をクリック

### 3. VBAエディターを開く

1. `Alt + F11` を押してVBAエディターを開く
2. または「開発」タブ →「Visual Basic」をクリック

### 4. 標準モジュールを追加

#### modWbsCommands モジュール

1. VBAエディターで「挿入」→「標準モジュール」を選択
2. 新規モジュールの名前を `modWbsCommands` に変更
3. `docs/vba/modWbsCommands.bas` の内容をコピー＆ペースト

```vba
' docs/vba/modWbsCommands.bas の内容を貼り付け
' - MoveTaskRowUp: 選択行を上に移動
' - MoveTaskRowDown: 選択行を下に移動
' - DuplicateTemplateSheet: Templateをコピーして新規PRJシート作成
' - UpdateTaskStatusFromKanban: カンバンからステータス更新
```

#### modProtection モジュール

1. 「挿入」→「標準モジュール」を選択
2. 新規モジュールの名前を `modProtection` に変更
3. `docs/vba/modProtection.bas` の内容をコピー＆ペースト

```vba
' docs/vba/modProtection.bas の内容を貼り付け
' - UnprotectAllSheets: 全シート保護解除
' - ProtectAllSheets: 全シート保護適用
' - ReapplyProtection: 保護の再適用
```

### 5. ThisWorkbook にコードを追加

1. VBAエディターのプロジェクトエクスプローラーで「ThisWorkbook」をダブルクリック
2. `docs/vba/ThisWorkbook.bas` の内容をコピー＆ペースト

```vba
' docs/vba/ThisWorkbook.bas の内容を貼り付け
' - NextProjectSheetName: 次のPRJ番号を採番
' - Workbook_Open: 起動時に保護を再適用
```

### 6. Kanban_View シートモジュールを編集（オプション）

カンバンカードのダブルクリックでステータス更新機能を有効にする場合:

1. プロジェクトエクスプローラーで Kanban_View シートをダブルクリック
2. `docs/vba/Kanban_View.bas` の内容をコピー＆ペースト

```vba
' docs/vba/Kanban_View.bas の内容を貼り付け
' - Worksheet_BeforeDoubleClick: カードダブルクリックでステータス遷移
```

### 7. 保存

1. `Ctrl + S` でブックを保存
2. 「マクロ有効ブック (.xlsm)」形式で保存されていることを確認

## ボタンとマクロの対応

| ボタン | マクロ名 | 機能 |
|--------|----------|------|
| ▲ Up | `modWbsCommands.MoveTaskRowUp` | 選択行を1行上に移動 |
| ▼ Down | `modWbsCommands.MoveTaskRowDown` | 選択行を1行下に移動 |
| 複製 | `modWbsCommands.DuplicateTemplateSheet` | Templateをコピーして新規PRJシート作成 |

## トラブルシューティング

### ボタンが機能しない

1. マクロが有効になっているか確認
2. VBAモジュールが正しくインポートされているか確認
3. モジュール名が正確か確認（大文字小文字を含む）

### 「マクロ 'xxx' が見つかりません」エラー

ボタンに割り当てられたマクロ名とVBAモジュール内の関数名が一致しているか確認:
- `modWbsCommands.MoveTaskRowUp`
- `modWbsCommands.MoveTaskRowDown`
- `modWbsCommands.DuplicateTemplateSheet`

### シート保護でエラーが発生する

`modProtection.UnprotectAllSheets` を実行してから操作を試してください。

## 注意事項

- VBAコードを追加した後、ファイルを保存して再度開くと正常に動作します
- セキュリティ設定によってはマクロの実行が制限される場合があります
- 会社のセキュリティポリシーに従ってマクロを有効化してください
