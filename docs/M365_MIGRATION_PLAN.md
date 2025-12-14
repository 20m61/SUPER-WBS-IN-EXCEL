# M365 完全移行計画

## 概要

Modern Excel PMS を Microsoft 365 専用として再設計し、動的配列関数（FILTER/LET/MAP/SEQUENCE/SORT/UNIQUE）をフル活用する。VBA依存を排除し、SharePoint / Excel Online での完全動作を保証する。

---

## 移行方針

### 廃止するもの
- 通常版（COUNTIF簡略化版）
- VBA版（.xlsm）
- 互換性のための機能制限

### 採用するもの
- FILTER/SORT/UNIQUE によるデータ抽出
- LET/LAMBDA による複雑な計算
- SEQUENCE による日付シーケンス生成
- MAP/REDUCE による配列処理
- XLOOKUP による高度な参照

---

## 機能別実装計画

### 1. Kanban_View（カンバンボード）

**現状**: 件数表示のみ → **目標**: 詳細カード表示（スピル）

```
カード形式:
┌─────────────────────┐
│ タスク名            │
│ 👤 担当者           │
│ 📅 期限: 2024/05/15 │
│ 📊 進捗: 50%        │
└─────────────────────┘
```

**実装数式** (LET + FILTER + TEXTJOIN):
```excel
=LET(
  _sheet, $B$2,
  _tasks, INDIRECT("'"&_sheet&"'!B5:B104"),
  _owners, INDIRECT("'"&_sheet&"'!C5:C104"),
  _ends, INDIRECT("'"&_sheet&"'!F5:F104"),
  _progress, INDIRECT("'"&_sheet&"'!G5:G104"),
  _status, INDIRECT("'"&_sheet&"'!H5:H104"),
  _filtered, FILTER(
    _tasks & CHAR(10) &
    "👤 " & _owners & CHAR(10) &
    "📅 " & TEXT(_ends,"yyyy/mm/dd") & CHAR(10) &
    "📊 " & TEXT(_progress,"0%"),
    _status="未着手",
    ""
  ),
  _filtered
)
```

### 2. Case_Master（案件管理）

**現状**: 件数表示のみ → **目標**: 施策一覧のスピル表示

**実装数式**:
```excel
=IF($H$1="", "← 案件IDを選択",
  IFERROR(
    FILTER(
      Measure_Master!A2:H104,
      Measure_Master!B2:B104=$H$1,
      "該当する施策がありません"
    ),
    ""
  )
)
```

### 3. Timeline/Gantt（タイムライン）

**現状**: 個別数式 → **目標**: SEQUENCEによる動的生成

**ヘッダー日付生成**:
```excel
=SEQUENCE(1, 60, $K$2, 1)
```

**ガントバー描画** (条件付き書式):
```excel
=AND(
  K$3 >= $D5,
  K$3 <= $F5,
  $D5 <> ""
)
```

### 4. 進捗サマリー

**新機能**: ダッシュボード的なサマリーセクション

```excel
=LET(
  _eff, E5:E104,
  _prg, G5:G104,
  _valid, FILTER(_eff, _eff<>""),
  _total, SUM(_valid),
  _completed, SUMPRODUCT(_eff, _prg),
  IF(_total=0, 0, _completed/_total)
)
```

### 5. 担当者別集計

**新機能**: UNIQUE + SUMIFS による自動集計

```excel
=LET(
  _owners, UNIQUE(FILTER(C5:C104, C5:C104<>"")),
  _efforts, MAP(_owners, LAMBDA(o, SUMIF(C5:C104, o, E5:E104))),
  HSTACK(_owners, _efforts)
)
```

---

## Python統合

### excel_generator.py の役割

1. **ブック生成**: OpenXML直接生成（現行）
2. **データ投入**: サンプルデータ/実データの自動投入
3. **レポート生成**: 進捗レポートのMarkdown/PDF出力
4. **バリデーション**: 生成ファイルの自動検証

### 新規: data_sync.py

SharePoint/OneDrive との連携を想定した同期スクリプト:

```python
# 構想
from office365.sharepoint.client_context import ClientContext

def upload_to_sharepoint(local_path, site_url, folder_path):
    """生成したブックをSharePointにアップロード"""
    pass

def download_and_update(site_url, file_path):
    """SharePointからダウンロードして更新"""
    pass
```

### 新規: report_generator.py

```python
def generate_progress_report(xlsx_path: Path) -> str:
    """ブックから進捗レポートを自動生成"""
    # openpyxlで読み込み
    # マークダウン形式でレポート出力
    pass
```

---

## UI/UX 強化

### 配色統一（Non-Excel Look）

| 要素 | 色コード | 用途 |
|------|----------|------|
| ヘッダー背景 | #2C3E50 | 全ヘッダー |
| 入力セル | #EAF2F8 | 編集可能セル |
| 計算セル | #F5F5F5 | 読み取り専用 |
| 完了 | #27AE60 | ステータス |
| 進行中 | #3498DB | ステータス |
| 遅延 | #E74C3C | ステータス |
| 未着手 | #95A5A6 | ステータス |

### アイコン活用

- 📋 リスト/一覧
- 📊 チャート/進捗
- 📅 日付/期限
- 👤 担当者
- ✅ 完了
- 🔄 進行中
- ⚠️ 遅延/警告
- 💡 ヒント/ガイド

### 操作ガイドの統一

全シートに操作ガイド行を設置:
```
💡 B2で WBS シートを選択すると、タスクカードが自動表示されます
```

---

## 実装順序

### Phase 1: コア機能（即時）
1. ✅ Kanban詳細カード（FILTER/LET）
2. ✅ Case_Masterドリルダウン（FILTER）
3. [ ] Timeline SEQUENCE化
4. [ ] 進捗サマリー追加

### Phase 2: 拡張機能
5. [ ] 担当者別集計
6. [ ] ステータス別サマリー
7. [ ] 警告条件付き書式の強化

### Phase 3: Python統合
8. [ ] レポート自動生成
9. [ ] SharePoint連携（構想）
10. [ ] CI/CD パイプライン

---

## 成功基準

1. **Kanban**: WBS選択時、全タスクが「名前+担当+期限+進捗」形式でスピル表示
2. **Case_Master**: 案件選択時、関連施策がすべてスピル表示
3. **Timeline**: SEQUENCE で60日分の日付が自動生成
4. **SharePoint**: Excel Online で全機能が動作
5. **テスト**: 自動テスト全項目パス

---

## ファイル構成（移行後）

```
output/
└── ModernExcelPMS.xlsx    # M365専用版（これが標準）

# 廃止
# - ModernExcelPMS_with_vba.xlsm
# - 旧互換版
```

---

## 備考

- VBA機能（行移動、シート複製）は手動操作で代替
- シート複製: 右クリック → コピー → 名前変更
- 行移動: 行選択 → 切り取り → 挿入
- ダブルクリック更新: 手動でステータス変更

SharePoint環境では上記の手動操作で十分対応可能。
