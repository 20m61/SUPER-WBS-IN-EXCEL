"""
Modern Excel PMS の雛形ブックを自動生成するスクリプト。
外部ライブラリに依存せず、OpenXML を直接書き出して `ModernExcelPMS.xlsm` を生成する。
"""

from __future__ import annotations

from dataclasses import dataclass
import argparse
from datetime import datetime
import os
import sys
from pathlib import Path
from typing import List, Mapping, Sequence, Tuple
from xml.sax.saxutils import escape
import zipfile

# toolsディレクトリをパスに追加
TOOLS_DIR = Path(__file__).resolve().parent
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

OUTPUT_PATH = Path(__file__).resolve().parent.parent / "ModernExcelPMS.xlsm"
VBA_SOURCE_DIR = Path(__file__).resolve().parent.parent / "docs" / "vba"

# シート保護パスワード（環境変数 PMS_SHEET_PASSWORD で上書き可能）
DEFAULT_SHEET_PASSWORD = "pms-2024"


def get_sheet_password() -> str:
    """環境変数からパスワードを取得、未設定時はデフォルトを返す。"""
    return os.environ.get("PMS_SHEET_PASSWORD", DEFAULT_SHEET_PASSWORD)


def excel_password_hash(password: str) -> str:
    """Excel 互換のパスワードハッシュを計算する（XOR ベース）。

    Excel 2003 以前のレガシー形式。sheetProtection の password 属性に使用。
    """
    if not password:
        return ""

    # Excel password hash algorithm
    pwd_hash = 0
    pwd_len = len(password)

    for i, char in enumerate(password):
        char_val = ord(char) << (i + 1)
        # Rotate left
        char_val = ((char_val >> 15) & 1) | ((char_val << 1) & 0x7FFF)
        pwd_hash ^= char_val

    pwd_hash ^= pwd_len
    pwd_hash ^= 0xCE4B

    return format(pwd_hash, "X")

# 共有データセット（レポート生成とシート生成の両方で再利用する）
HOLIDAYS = ["2024-01-01", "2024-02-12", "2024-04-29", "2024-05-03", "2024-05-04", "2024-05-05"]
MEMBERS = ["PM_佐藤", "TL_田中", "DEV_鈴木", "EXT_山田", "EXT_山田"]
STATUSES = ["未着手", "進行中", "遅延", "完了"]
CASES = [("CASE-001", "新規システム導入プロジェクト")]
MEASURES = [
    ("ME-001", "CASE-001", "要件定義・設計フェーズ", "2025-12-15", "PRJ_001"),
    ("ME-002", "CASE-001", "開発・テストフェーズ", "2026-01-06", "PRJ_002"),
]


@dataclass
class SampleTask:
    """サンプルタスクのデータを保持する。"""

    lv: int
    name: str
    owner: str
    start_date: str
    effort: int
    progress: float  # 0.0 〜 1.0

    @property
    def status(self) -> str:
        """進捗率からステータスを判定する（簡易版）。"""
        if self.progress >= 1.0:
            return "完了"
        elif self.progress > 0:
            return "進行中"
        else:
            return "未着手"


# サンプルタスクデータ（PRJ_001 に配置）
# 新規システム導入推進プロジェクト Phase0: 基盤再構築
SAMPLE_TASKS: List[SampleTask] = [
    # Phase 0: 基盤再構築 (12/15-12/27)
    SampleTask(lv=1, name="要件定義・設計フェーズ", owner="PM_佐藤", start_date="2025-12-15", effort=10, progress=0.0),
    SampleTask(lv=2, name="ヒアリング実施", owner="PM_佐藤", start_date="2025-12-15", effort=4, progress=0.0),
    SampleTask(lv=2, name="要件定義書作成", owner="PM_佐藤", start_date="2025-12-20", effort=5, progress=0.0),
    SampleTask(lv=2, name="要件レビュー", owner="TL_田中", start_date="2025-12-15", effort=8, progress=0.0),
    SampleTask(lv=2, name="スコープ確定", owner="PM_佐藤", start_date="2025-12-15", effort=9, progress=0.0),
    SampleTask(lv=2, name="WBS作成", owner="DEV_鈴木", start_date="2025-12-15", effort=9, progress=0.0),
    # Phase 1: 集中導入 (1/6-1/20)
    SampleTask(lv=1, name="開発・テストフェーズ", owner="PM_佐藤", start_date="2026-01-06", effort=15, progress=0.0),
    SampleTask(lv=2, name="基本設計(外部ベンダー)", owner="EXT_山田", start_date="2025-12-25", effort=12, progress=0.0),
    SampleTask(lv=2, name="設計レビュー", owner="PM_佐藤", start_date="2026-01-10", effort=1, progress=0.0),
    SampleTask(lv=2, name="詳細設計", owner="PM_佐藤", start_date="2026-01-10", effort=5, progress=0.0),
    SampleTask(lv=2, name="テスト計画作成", owner="DEV_鈴木", start_date="2026-01-10", effort=10, progress=0.0),
    SampleTask(lv=2, name="環境構築", owner="PM_佐藤", start_date="2026-01-15", effort=10, progress=0.0),
    # Phase 2: 横展開・測定 (2/1-2/25)
    SampleTask(lv=1, name="開発フェーズ", owner="TL_田中", start_date="2026-02-01", effort=20, progress=0.0),
    SampleTask(lv=2, name="機能実装", owner="TL_田中", start_date="2026-01-20", effort=20, progress=0.0),
    SampleTask(lv=2, name="単体テスト", owner="PM_佐藤", start_date="2026-02-10", effort=10, progress=0.0),
    SampleTask(lv=2, name="結合テスト", owner="TL_田中", start_date="2026-02-01", effort=10, progress=0.0),
    SampleTask(lv=2, name="ドキュメント作成", owner="TL_田中", start_date="2026-02-10", effort=8, progress=0.0),
    SampleTask(lv=2, name="リリース準備", owner="PM_佐藤", start_date="2026-02-25", effort=1, progress=0.0),
]


def date_to_excel_serial(date_str: str) -> int:
    """日付文字列をExcelシリアル値に変換する。

    Excelでは1900年1月1日を1とするシリアル値を使用。
    ただし1900年2月29日のバグがあるため、1900年3月1日以降は+1する。
    """
    from datetime import datetime
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    # Excel epoch: 1899-12-30 (Excelの1900年バグを考慮)
    excel_epoch = datetime(1899, 12, 30)
    delta = dt - excel_epoch
    return delta.days


def calculate_weighted_progress(tasks: List[SampleTask]) -> float:
    """工数加重平均で進捗率を計算する。"""
    total_effort = sum(t.effort for t in tasks)
    if total_effort == 0:
        return 0.0
    return sum(t.effort * t.progress for t in tasks) / total_effort


def count_by_status(tasks: List[SampleTask]) -> Mapping[str, int]:
    """ステータス別のタスク数を集計する。"""
    counts: dict[str, int] = {s: 0 for s in STATUSES}
    for task in tasks:
        status = task.status
        if status in counts:
            counts[status] += 1
    return counts


@dataclass
class Formula:
    """セルに設定する数式を保持する。"""

    expr: str

    def __post_init__(self) -> None:
        if self.expr.startswith("="):
            self.expr = self.expr[1:]


def col_letter(index: int) -> str:
    """列番号を Excel の列名に変換する。"""
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def cell_ref(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


def cell_xml(row: int, col: int, value, style_id: int = 0) -> str:
    """セルの XML を生成する。

    Args:
        style_id: 0=ロック（デフォルト）、1=ロック解除

    空文字列の場合はスタイルのみ適用（値なし）。
    """
    ref = cell_ref(row, col)
    style_attr = f' s="{style_id}"' if style_id else ""
    if isinstance(value, Formula):
        return f"<c r=\"{ref}\"{style_attr}><f>{escape(value.expr)}</f></c>"
    if isinstance(value, str):
        # 空文字列の場合はスタイルのみ適用（値なし）
        if value == "":
            return f"<c r=\"{ref}\"{style_attr}/>"
        return f"<c r=\"{ref}\"{style_attr} t=\"inlineStr\"><is><t>{escape(value)}</t></is></c>"
    if value is None:
        return ""
    return f"<c r=\"{ref}\"{style_attr}><v>{value}</v></c>"


@dataclass
class SheetProtection:
    """シート保護の設定を保持する。"""

    password_hash: str = ""
    allow_insert_rows: bool = False

    def to_xml(self) -> str:
        """<sheetProtection> 要素を生成する。"""
        attrs = [
            'sheet="1"',
            'objects="1"',
            'formatCells="0"',  # 0=許可
            'sort="0"',
            'autoFilter="0"',
        ]
        if self.password_hash:
            attrs.append(f'password="{self.password_hash}"')
        if self.allow_insert_rows:
            attrs.append('insertRows="0"')  # 0=許可
        else:
            attrs.append('insertRows="1"')  # 1=禁止
        return f"<sheetProtection {' '.join(attrs)}/>"


# XML 宣言（全 XML ファイルの先頭に付与）
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


@dataclass
class ColumnDef:
    """列幅の定義。"""
    min_col: int  # 1-indexed
    max_col: int  # 1-indexed
    width: float
    custom_width: bool = True


def cols_xml(col_defs: Sequence[ColumnDef]) -> str:
    """<cols> 要素を生成する。"""
    if not col_defs:
        return ""
    cols = []
    for cd in col_defs:
        cols.append(
            f'<col min="{cd.min_col}" max="{cd.max_col}" width="{cd.width}" '
            f'customWidth="{1 if cd.custom_width else 0}"/>'
        )
    return "<cols>" + "".join(cols) + "</cols>"


def sheet_views_xml(
    freeze_row: int = 0,
    freeze_col: int = 0,
    active_cell: str = "A1",
    tab_selected: bool = False,
    show_grid_lines: bool = False,
) -> str:
    """<sheetViews> 要素を生成する。

    Args:
        freeze_row: フリーズする行数
        freeze_col: フリーズする列数
        active_cell: アクティブセル
        tab_selected: タブ選択状態
        show_grid_lines: グリッド線表示（デフォルト: False - Non-Excel Look）
    """
    pane = ""
    if freeze_row > 0 or freeze_col > 0:
        top_left = cell_ref(freeze_row + 1, freeze_col + 1) if freeze_row > 0 or freeze_col > 0 else "A1"
        x_split = f' xSplit="{freeze_col}"' if freeze_col > 0 else ""
        y_split = f' ySplit="{freeze_row}"' if freeze_row > 0 else ""
        pane = f'<pane{x_split}{y_split} topLeftCell="{top_left}" activePane="bottomRight" state="frozen"/>'

    selected = ' tabSelected="1"' if tab_selected else ""
    grid_attr = '' if show_grid_lines else ' showGridLines="0"'
    return (
        "<sheetViews>"
        f'<sheetView workbookViewId="0"{selected}{grid_attr}>'
        f'{pane}'
        f'<selection activeCell="{active_cell}" sqref="{active_cell}"/>'
        "</sheetView>"
        "</sheetViews>"
    )


def worksheet_xml(
    cells: Sequence[Tuple[int, int, object]],
    data_validations: str | None = None,
    conditional_formattings: Sequence[str] | None = None,
    sheet_protection: SheetProtection | None = None,
    unlocked_cells: set[Tuple[int, int]] | None = None,
    legacy_drawing_rid: str | None = None,
    column_defs: Sequence[ColumnDef] | None = None,
    freeze_row: int = 0,
    freeze_col: int = 0,
    cell_styles: dict[Tuple[int, int], int] | None = None,
    show_grid_lines: bool = False,
) -> str:
    """ワークシート XML を生成する。

    Args:
        cells: (row, col, value) のセルデータ
        data_validations: データ検証 XML
        conditional_formattings: 条件付き書式 XML リスト
        sheet_protection: シート保護設定
        unlocked_cells: ロック解除するセルの (row, col) セット（旧方式、cell_styles優先）
        legacy_drawing_rid: VML描画への参照ID（ボタン用）
        column_defs: 列幅定義のリスト
        freeze_row: フリーズする行数（ヘッダー固定用）
        freeze_col: フリーズする列数
        cell_styles: セル座標からスタイルIDへのマッピング
        show_grid_lines: グリッド線表示（デフォルト: False - Non-Excel Look）
    """
    rows = {}
    for row, col, value in cells:
        rows.setdefault(row, {})[col] = value

    unlocked = unlocked_cells or set()
    styles_map = cell_styles or {}

    xml_lines: List[str] = [
        XML_DECL,
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]

    # シートビュー（フリーズペイン、グリッド線非表示）- 常に出力
    xml_lines.append(sheet_views_xml(freeze_row, freeze_col, show_grid_lines=show_grid_lines))

    # 列幅定義
    if column_defs:
        xml_lines.append(cols_xml(column_defs))

    xml_lines.append("<sheetData>")

    for row_idx in sorted(rows):
        xml_lines.append(f'<row r="{row_idx}">')
        for col_idx in sorted(rows[row_idx]):
            # スタイルマップがあればそれを優先、なければ unlocked_cells で判定
            if (row_idx, col_idx) in styles_map:
                style_id = styles_map[(row_idx, col_idx)]
            elif (row_idx, col_idx) in unlocked:
                style_id = STYLE_UNLOCKED
            else:
                style_id = STYLE_LOCKED
            xml_lines.append(cell_xml(row_idx, col_idx, rows[row_idx][col_idx], style_id))
        xml_lines.append("</row>")

    xml_lines.append("</sheetData>")

    # OpenXML 仕様に従った要素順序:
    # sheetData → sheetProtection → conditionalFormatting → dataValidations → legacyDrawing
    if sheet_protection:
        xml_lines.append(sheet_protection.to_xml())

    if conditional_formattings:
        xml_lines.extend(conditional_formattings)

    if data_validations:
        xml_lines.append(data_validations)

    # VML描画（ボタン）への参照
    if legacy_drawing_rid:
        xml_lines.append(f'<legacyDrawing r:id="{legacy_drawing_rid}"/>')

    xml_lines.append("</worksheet>")
    return "".join(xml_lines)


@dataclass
class ButtonDefinition:
    """ボタンの定義を保持する。"""

    name: str
    macro_name: str
    row: int  # 0-indexed
    col: int  # 0-indexed
    width: int = 80  # pixels
    height: int = 24  # pixels
    text: str = ""


def vml_drawing_xml(buttons: Sequence[ButtonDefinition], sheet_name: str) -> str:
    """VML形式のボタン描画XMLを生成する。

    Excel Form Controls はVML形式で定義される。
    """
    shapes = []
    for idx, btn in enumerate(buttons, start=1):
        # VMLの座標系: 列と行を指定、オフセットはピクセル単位
        left_col = btn.col
        top_row = btn.row
        right_col = btn.col + 1
        bottom_row = btn.row + 1

        shape = f'''<v:shape id="_x0000_s{1024 + idx}" type="#_x0000_t201"
 style="position:absolute;margin-left:6pt;margin-top:3pt;width:{btn.width}pt;height:{btn.height}pt;z-index:{idx}"
 o:button="t" fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto">
 <v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>
 <v:textbox style="mso-direction-alt:auto" o:singleclick="f">
  <div style="text-align:center"><font face="Meiryo UI" size="160" color="#000000">{escape(btn.text or btn.name)}</font></div>
 </v:textbox>
 <x:ClientData ObjectType="Button">
  <x:Anchor>{left_col}, 8, {top_row}, 6, {right_col}, 72, {bottom_row}, 2</x:Anchor>
  <x:PrintObject>False</x:PrintObject>
  <x:AutoFill>False</x:AutoFill>
  <x:FmlaMacro>{escape(btn.macro_name)}</x:FmlaMacro>
  <x:TextHAlign>Center</x:TextHAlign>
  <x:TextVAlign>Center</x:TextVAlign>
 </x:ClientData>
</v:shape>'''
        shapes.append(shape)

    return f'''<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout>
 <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype>
{"".join(shapes)}
</xml>'''


def worksheet_rels_xml(vml_rid: str | None = None, vml_filename: str = "vmlDrawing1.vml") -> str | None:
    """ワークシートのリレーションシップXMLを生成する。

    Args:
        vml_rid: VML描画への参照ID
        vml_filename: VMLファイル名（xl/drawings/以下）
    """
    if not vml_rid:
        return None
    return (
        XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="{vml_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/{vml_filename}"/>'
        "</Relationships>"
    )


def content_types_xml(sheet_count: int, has_vml: bool = False, has_vba: bool = False) -> str:
    """[Content_Types].xml を生成する。

    Args:
        sheet_count: シート数
        has_vml: VML 描画を含むか
        has_vba: VBA プロジェクトを含むか
    """
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for idx in range(1, sheet_count + 1)
    )
    vml_default = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' if has_vml else ""

    # マクロ有効ブック (.xlsm) か通常ブック (.xlsx) かでコンテンツタイプを切り替え
    if has_vba:
        workbook_content_type = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        vba_override = '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
    else:
        workbook_content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        vba_override = ""

    return (
        XML_DECL +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{vml_default}"
        f'<Override PartName="/xl/workbook.xml" ContentType="{workbook_content_type}"/>'
        f"{overrides}"
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f"{vba_override}"
        "</Types>"
    )


def root_rels_xml() -> str:
    return (
        XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def workbook_xml(sheet_names: Sequence[str], defined_names: Mapping[str, str] | None = None) -> str:
    sheets_xml = "".join(
        f'<sheet name="{escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        for idx, name in enumerate(sheet_names, start=1)
    )

    defined_names_xml = ""
    if defined_names:
        defined_names_xml = "<definedNames>" + "".join(
            f'<definedName name="{escape(name)}">{escape(ref)}</definedName>'
            for name, ref in defined_names.items()
        ) + "</definedNames>"

    return (
        XML_DECL +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets_xml}</sheets>"
        f"{defined_names_xml}"
        "</workbook>"
    )


def workbook_rels_xml(sheet_count: int, has_vba: bool = False) -> str:
    """xl/_rels/workbook.xml.rels を生成する。"""
    rels = "".join(
        f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
        for idx in range(1, sheet_count + 1)
    )
    rels += (
        f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    if has_vba:
        rels += (
            f'<Relationship Id="rId{sheet_count + 2}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
        )
    return XML_DECL + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>'


def styles_xml() -> str:
    """スタイルシートを生成する。

    セルスタイル:
      - xfId=0: 標準（ロック）
      - xfId=1: ロック解除（編集可能セル用）
      - xfId=2: ヘッダー（青背景・白太字）
      - xfId=3: 入力セル（薄青背景・罫線）
      - xfId=4: タイトル（大きい太字）
      - xfId=5: 計算セル（グレー背景・読み取り専用）
      - xfId=6: サブヘッダー（薄青背景・太字）
      - xfId=7: 日付ヘッダー（小さいフォント・中央揃え）
      - xfId=8: 説明テキスト（イタリック）
      - xfId=9: パーセント表示（入力セル）
    """
    # フォント定義
    fonts = (
        '<fonts count="6">'
        '<font><sz val="11"/><color theme="1"/><name val="Meiryo UI"/><family val="2"/></font>'  # 0: 標準
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/></font>'  # 1: ヘッダー用（白太字）
        '<font><b/><sz val="14"/><color theme="1"/><name val="Meiryo UI"/><family val="2"/></font>'  # 2: タイトル用
        '<font><b/><sz val="11"/><color theme="1"/><name val="Meiryo UI"/><family val="2"/></font>'  # 3: サブヘッダー用（太字）
        '<font><sz val="9"/><color theme="1"/><name val="Meiryo UI"/><family val="2"/></font>'  # 4: 小さいフォント
        '<font><i/><sz val="10"/><color rgb="FF666666"/><name val="Meiryo UI"/><family val="2"/></font>'  # 5: 説明用（イタリック・グレー）
        '</fonts>'
    )

    # 塗りつぶし定義
    fills = (
        '<fills count="6">'
        '<fill><patternFill patternType="none"/></fill>'  # 0: なし
        '<fill><patternFill patternType="gray125"/></fill>'  # 1: グレーパターン
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2C3E50"/><bgColor indexed="64"/></patternFill></fill>'  # 2: ダークブルー（ヘッダー）
        '<fill><patternFill patternType="solid"><fgColor rgb="FFEAF2F8"/><bgColor indexed="64"/></patternFill></fill>'  # 3: 薄青（入力セル）
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF5F5F5"/><bgColor indexed="64"/></patternFill></fill>'  # 4: 薄グレー（計算セル）
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD5E8F7"/><bgColor indexed="64"/></patternFill></fill>'  # 5: 薄青（サブヘッダー）
        '</fills>'
    )

    # 罫線定義
    borders = (
        '<borders count="4">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'  # 0: なし
        '<border>'  # 1: 薄い罫線（全方向）
        '<left style="thin"><color indexed="64"/></left>'
        '<right style="thin"><color indexed="64"/></right>'
        '<top style="thin"><color indexed="64"/></top>'
        '<bottom style="thin"><color indexed="64"/></bottom>'
        '<diagonal/>'
        '</border>'
        '<border>'  # 2: 下線のみ
        '<left/><right/><top/>'
        '<bottom style="thin"><color indexed="64"/></bottom>'
        '<diagonal/>'
        '</border>'
        '<border>'  # 3: 太い下線
        '<left/><right/><top/>'
        '<bottom style="medium"><color rgb="FF2C3E50"/></bottom>'
        '<diagonal/>'
        '</border>'
        '</borders>'
    )

    # 数値フォーマット
    num_fmts = (
        '<numFmts count="2">'
        '<numFmt numFmtId="164" formatCode="yyyy/mm/dd"/>'  # 日付
        '<numFmt numFmtId="165" formatCode="0%"/>'  # パーセント
        '</numFmts>'
    )

    # セルフォーマット定義
    cell_xfs = (
        '<cellXfs count="11">'
        # 0: 標準（ロック）
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        # 1: ロック解除（編集可能）
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>'
        # 2: ヘッダー（青背景・白太字・罫線・中央揃え）
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        # 3: 入力セル（薄青背景・罫線・ロック解除）
        '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyProtection="1"><protection locked="0"/></xf>'
        # 4: タイトル（大きい太字）
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        # 5: 計算セル（グレー背景・罫線）
        '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>'
        # 6: サブヘッダー（薄青背景・太字・罫線）
        '<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        # 7: 日付ヘッダー（小フォント・中央揃え・下線）
        '<xf numFmtId="164" fontId="4" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
        # 8: 説明テキスト（イタリック）
        '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        # 9: パーセント入力セル
        '<xf numFmtId="165" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyProtection="1"><protection locked="0"/></xf>'
        # 10: 日付入力セル（薄青背景・日付フォーマット・罫線・ロック解除）
        '<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyProtection="1"><protection locked="0"/></xf>'
        '</cellXfs>'
    )

    # 条件付き書式用スタイル（dxf）
    dxfs = (
        '<dxfs count="11">'
        '<dxf><border><right style="medium"><color rgb="FFE74C3C"/></right></border></dxf>'  # 0: 今日ライン
        '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF95A5A6"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 1: 完了（グレー）
        '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFE74C3C"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 2: 遅延（赤）
        '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF3498DB"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 3: 進行中（青）
        '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFECF0F1"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 4: 未着手（薄グレー）
        '<dxf><font><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="FF3498DB"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 5: 進行中ステータス
        '<dxf><font><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFE74C3C"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 6: 遅延ステータス
        '<dxf><font><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="FF27AE60"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 7: 完了ステータス
        '<dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="FF34495E"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 8: Lv1行（濃紺背景・白太字）
        '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFF39C12"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 9: 警告（オレンジ）- 未リンク
        '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFF1C40F"/><bgColor indexed="64"/></patternFill></fill></dxf>'  # 10: 警告（黄）- 範囲外
        '</dxfs>'
    )

    return (
        XML_DECL +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'{num_fmts}'
        f'{fonts}'
        f'{fills}'
        f'{borders}'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'{cell_xfs}'
        '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>'
        f'{dxfs}'
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>'
        '</styleSheet>'
    )


# スタイル ID 定数（styles_xml の cellXfs と対応）
STYLE_LOCKED = 0
STYLE_UNLOCKED = 1
STYLE_HEADER = 2
STYLE_INPUT = 3
STYLE_TITLE = 4
STYLE_CALC = 5
STYLE_SUBHEADER = 6
STYLE_DATE_HEADER = 7
STYLE_DESCRIPTION = 8
STYLE_PERCENT_INPUT = 9
STYLE_DATE_INPUT = 10


def load_vba_modules() -> Mapping[str, str]:
    modules: dict[str, str] = {}
    if not VBA_SOURCE_DIR.exists():
        return modules

    for path in sorted(VBA_SOURCE_DIR.glob("*.bas")):
        modules[path.stem] = path.read_text(encoding="utf-8")

    for path in sorted(VBA_SOURCE_DIR.glob("*.cls")):
        modules[path.stem] = path.read_text(encoding="utf-8")

    return modules


def vba_project_binary(modules: Mapping[str, str], regenerate: bool = False) -> bytes | None:
    """VBAプロジェクトバイナリを取得または生成する。

    vbaProject.binはOLE複合ドキュメント形式である必要がある。

    1. regenerate=Falseかつテンプレートファイルが存在する場合はそれを使用
    2. テンプレートがない、またはregenerate=Trueの場合は自動生成を試みる
    3. 自動生成に失敗した場合はNoneを返す

    Args:
        modules: VBAモジュールの辞書 {モジュール名: コード}
        regenerate: 既存のキャッシュを無視して再生成するか

    Returns:
        バイナリデータ、または生成に失敗した場合はNone
    """
    # テンプレートファイルのパス
    template_path = VBA_SOURCE_DIR / "vbaProject.bin"

    if not regenerate and template_path.exists():
        print(f"📦 vbaProject.bin を読み込み: {template_path}")
        return template_path.read_bytes()

    # 自動生成を試みる
    if regenerate:
        print("🔄 vbaProject.bin を再生成中...")
    else:
        print("🔧 vbaProject.bin を自動生成中...")

    try:
        from create_vba_binary import generate_vba_project_bin
        vba_binary = generate_vba_project_bin(dict(modules))
        # 生成したバイナリをキャッシュとして保存
        template_path.write_bytes(vba_binary)
        print(f"✅ vbaProject.bin を生成しました ({len(vba_binary)} bytes)")
        return vba_binary
    except ImportError:
        # create_vba_binary モジュールが見つからない場合
        print("⚠️  警告: create_vba_binary モジュールが見つかりません")
        print("   VBA機能は手動で追加する必要があります")
        return None
    except Exception as e:
        print(f"⚠️  警告: vbaProject.bin の生成に失敗しました: {e}")
        print("   VBA機能は手動で追加する必要があります")
        return None


# --------------------------- VBA モジュール計画 ---------------------------


@dataclass(frozen=True)
class VBAProcedurePlan:
    """VBA プロシージャの名称と役割をまとめる。"""

    name: str
    description: str


@dataclass(frozen=True)
class VBAModulePlan:
    """VBA モジュールの種類と配置方針を保持する。"""

    module_type: str  # Standard / Worksheet / ThisWorkbook
    module_name: str
    description: str
    procedures: List[VBAProcedurePlan]


# 後続の VBA 自動生成で参照するモジュール配置と主要プロシージャ
VBA_MODULE_PLAN: List[VBAModulePlan] = [
    VBAModulePlan(
        module_type="Standard",
        module_name="modWbsCommands",
        description="行入れ替えやテンプレート複製など、WBS シート共通のコマンド群を置く。",
        procedures=[
            VBAProcedurePlan(
                name="MoveTaskRowUp",
                description="選択行を一行上へスワップする。Up/Down ボタンのマクロ割当先。",
            ),
            VBAProcedurePlan(
                name="MoveTaskRowDown",
                description="選択行を一行下へスワップする。Up/Down ボタンのマクロ割当先。",
            ),
            VBAProcedurePlan(
                name="DuplicateTemplateSheet",
                description="Template を複製し、ThisWorkbook の採番関数から取得したシート名で貼り付ける。",
            ),
            VBAProcedurePlan(
                name="UpdateTaskStatusFromKanban",
                description="カンバンのセルから対象タスクを特定し、ステータスを書き換える共通処理。",
            ),
        ],
    ),
    VBAModulePlan(
        module_type="Standard",
        module_name="modProtection",
        description="シート保護の一括適用・解除を行い、マクロ操作時の保護エラーを防ぐ。",
        procedures=[
            VBAProcedurePlan(
                name="UnprotectAllSheets",
                description="全シートの保護をまとめて解除する。保護パスワードは定数で集中管理する。",
            ),
            VBAProcedurePlan(
                name="ProtectAllSheets",
                description="編集可能セルだけを解放した状態で保護をかけ直す。UserInterfaceOnly を True に設定してマクロ操作を許可。",
            ),
            VBAProcedurePlan(
                name="ReapplyProtection",
                description="解除→再保護を一括実行するラッパー。設定変更時の再適用に使う。",
            ),
        ],
    ),
    VBAModulePlan(
        module_type="Worksheet",
        module_name="Kanban_View",
        description="カンバンシートのイベント ハンドラを保持。ダブルクリックでステータス更新を呼び出す。",
        procedures=[
            VBAProcedurePlan(
                name="Worksheet_BeforeDoubleClick",
                description="カードセルのダブルクリックで UpdateTaskStatusFromKanban を呼び出し、イベントをキャンセルする。",
            ),
        ],
    ),
    VBAModulePlan(
        module_type="ThisWorkbook",
        module_name="ThisWorkbook",
        description="ブック全体で共有するユーティリティを定義。テンプレート複製時のシート名採番を行う。",
        procedures=[
            VBAProcedurePlan(
                name="NextProjectSheetName",
                description="既存の PRJ_xxx を走査し、次に付与する連番シート名を返す。",
            ),
        ],
    ),
]


# --------------------------- シート定義 ---------------------------

def config_sheet(password_hash: str = "") -> str:
    """Config シートを生成する。

    編集可能: 祝日 B4:B200、担当者 D4:D200、ステータス F4:F200
    """
    cells: List[Tuple[int, int, object]] = []
    styles: dict[Tuple[int, int], int] = {}

    # タイトル行
    cells.append((1, 1, "⚙️ 設定シート"))
    styles[(1, 1)] = STYLE_TITLE

    # 操作ガイド
    cells.append((2, 1, "💡 薄青のセルにデータを追加・編集できます。WBS シートで使用されます。"))
    styles[(2, 1)] = STYLE_DESCRIPTION

    # 祝日リスト
    cells.append((3, 1, "祝日リスト"))
    styles[(3, 1)] = STYLE_SUBHEADER
    cells.append((3, 2, "日付"))
    styles[(3, 2)] = STYLE_HEADER

    # 担当者リスト（マスタ）
    cells.append((3, 3, "担当者リスト"))
    styles[(3, 3)] = STYLE_SUBHEADER
    cells.append((3, 4, "担当者"))
    styles[(3, 4)] = STYLE_HEADER

    # ステータスリスト（候補）
    cells.append((3, 5, "ステータスリスト"))
    styles[(3, 5)] = STYLE_SUBHEADER
    cells.append((3, 6, "ステータス"))
    styles[(3, 6)] = STYLE_HEADER

    # データ行
    for idx, day in enumerate(HOLIDAYS, start=4):
        cells.append((idx, 2, day))
        styles[(idx, 2)] = STYLE_INPUT
    for idx, member in enumerate(MEMBERS, start=4):
        cells.append((idx, 4, member))
        styles[(idx, 4)] = STYLE_INPUT
    for idx, status in enumerate(STATUSES, start=4):
        cells.append((idx, 6, status))
        styles[(idx, 6)] = STYLE_INPUT

    # 空の入力行にもスタイルを設定
    for row in range(4 + len(HOLIDAYS), 21):
        styles[(row, 2)] = STYLE_INPUT
    for row in range(4 + len(MEMBERS), 21):
        styles[(row, 4)] = STYLE_INPUT
    for row in range(4 + len(STATUSES), 21):
        styles[(row, 6)] = STYLE_INPUT

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=True)
    return worksheet_xml(
        cells,
        sheet_protection=protection,
        column_defs=get_config_column_defs(),
        freeze_row=3,
        cell_styles=styles,
    )


@dataclass
class StyledCell:
    """スタイル付きセルを表現する。"""
    value: object
    style_id: int = STYLE_LOCKED


def template_cells(sample: bool = False) -> Tuple[List[Tuple[int, int, object]], dict[Tuple[int, int], int]]:
    """WBS シートのセルデータとスタイルマッピングを返す。"""
    cells: List[Tuple[int, int, object]] = []
    styles: dict[Tuple[int, int], int] = {}

    # タイトル行
    cells.append((1, 1, "プロジェクト名"))
    styles[(1, 1)] = STYLE_SUBHEADER
    cells.append((1, 2, ""))  # プロジェクト名入力欄
    styles[(1, 2)] = STYLE_INPUT

    # 操作ガイド
    cells.append((2, 1, "💡 薄青のセルに入力できます。ステータスは自動計算されます。"))
    styles[(2, 1)] = STYLE_DESCRIPTION

    # ヘッダー行
    headers = ["Lv", "タスク名", "担当", "開始日", "工数(日)", "終了日", "進捗率", "ステータス", "備考"]
    for col, header in enumerate(headers, start=1):
        cells.append((4, col, header))
        styles[(4, col)] = STYLE_HEADER

    # 全体進捗エリア
    cells.append((1, 10, "全体進捗"))
    styles[(1, 10)] = STYLE_SUBHEADER
    cells.append((2, 10, Formula("LET(_eff,E5:E104,_prg,G5:G104,_total,SUM(_eff),IF(_total=0,0,SUMPRODUCT(_eff,_prg)/_total))")))
    styles[(2, 10)] = STYLE_CALC

    # ガントチャートエリア
    cells.append((1, 11, "ガント開始日"))
    styles[(1, 11)] = STYLE_SUBHEADER
    # サンプルデータの最初の日付に合わせる（2024-12-09 = シリアル値45635）
    if sample:
        cells.append((2, 11, date_to_excel_serial("2024-12-09")))
    else:
        cells.append((2, 11, Formula("TODAY()")))
    styles[(2, 11)] = STYLE_DATE_INPUT

    # ガント日付ヘッダー（SEQUENCE使用 - M365専用）
    # K3に SEQUENCE(1, 60, $K$2, 1) を配置し、60日分の日付を動的生成
    gantt_start_col = 11
    gantt_columns = 60  # 約2ヶ月分
    # SEQUENCEでスピル表示（M365専用機能）
    cells.append((3, gantt_start_col, Formula('IF($K$2="","",SEQUENCE(1,60,$K$2,1))')))
    styles[(3, gantt_start_col)] = STYLE_DATE_HEADER
    # 残りの列にもスタイルを適用（スピル先）
    for offset in range(1, gantt_columns):
        styles[(3, gantt_start_col + offset)] = STYLE_DATE_HEADER

    # タスク行のテンプレート（数式のみ）- 十分な行数を確保
    task_rows = 20 if sample else 14
    for row in range(5, 5 + task_rows):
        # 終了日（計算列）
        cells.append((row, 6, Formula(f'IF(OR(D{row}="",E{row}="")," ",WORKDAY(D{row},E{row}-1,Config!$B$4:$B$20))')))
        styles[(row, 6)] = STYLE_CALC
        # ステータス（計算列）
        cells.append((row, 8, Formula(f'IFS(G{row}=1,"完了",AND(F{row}<TODAY(),G{row}<1),"遅延",AND(D{row}<=TODAY(),G{row}<1),"進行中",TRUE,"未着手")')))
        styles[(row, 8)] = STYLE_CALC

        # 入力セルのスタイルを設定（空の値でもセルを作成してスタイルを適用）
        for col in [1, 2, 3, 5, 9]:  # Lv, タスク名, 担当, 工数, 備考
            styles[(row, col)] = STYLE_INPUT
            # サンプルデータがない行には空文字を入れてスタイルを適用
            if not sample or row >= 5 + len(SAMPLE_TASKS):
                cells.append((row, col, ""))
        styles[(row, 4)] = STYLE_DATE_INPUT  # 開始日
        if not sample or row >= 5 + len(SAMPLE_TASKS):
            cells.append((row, 4, ""))  # 空の開始日セル
        styles[(row, 7)] = STYLE_PERCENT_INPUT  # 進捗率
        if not sample or row >= 5 + len(SAMPLE_TASKS):
            cells.append((row, 7, ""))  # 空の進捗率セル

    # サンプルデータ
    if sample:
        for row_offset, task in enumerate(SAMPLE_TASKS):
            row = 5 + row_offset
            # レベルに応じたインデント（Lv2は「  └ 」を付ける）
            indent = "  └ " if task.lv >= 2 else ""
            task_name = f"{indent}{task.name}"
            # 日付をExcelシリアル値に変換
            date_serial = date_to_excel_serial(task.start_date)
            cells.extend([
                (row, 1, task.lv),
                (row, 2, task_name),
                (row, 3, task.owner),
                (row, 4, date_serial),  # シリアル値で保存
                (row, 5, task.effort),
                (row, 7, task.progress),
            ])
            # 日付セルのスタイルを日付フォーマットに
            styles[(row, 4)] = STYLE_DATE_INPUT

    return cells, styles


def template_data_validations() -> str:
    """WBSシート用のデータバリデーションを生成する。

    バリデーション内容:
    - 担当者（C列）: Configシートのリストから選択
    - 開始日（D列）: 日付形式のみ許可
    - 工数（E列）: 1〜100の整数
    - 進捗率（G列）: 0〜1（0%〜100%）の範囲
    - ステータス（H列）: Configシートのリストから選択（入力禁止・数式専用）
    """
    return (
        '<dataValidations count="5">'
        # 担当者: リスト選択
        '<dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" '
        'promptTitle="担当者" prompt="Configシートで定義された担当者から選択してください" '
        'errorTitle="入力エラー" error="リストから選択してください" sqref="C5:C104">'
        "<formula1>Config!$D$4:$D$20</formula1>"
        "</dataValidation>"
        # 開始日: 日付形式
        '<dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
        'promptTitle="開始日" prompt="タスクの開始日を入力（例: 2024/05/01）" '
        'errorTitle="入力エラー" error="有効な日付を入力してください（例: 2024/05/01）" sqref="D5:D104">'
        "<formula1>1</formula1><formula2>109574</formula2>"  # 1900/1/1 〜 2199/12/31
        "</dataValidation>"
        # 工数: 1〜100の整数
        '<dataValidation type="whole" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
        'promptTitle="工数（日）" prompt="1〜100の整数を入力してください" '
        'errorTitle="入力エラー" error="工数は1〜100の整数で入力してください" sqref="E5:E104">'
        "<formula1>1</formula1><formula2>100</formula2>"
        "</dataValidation>"
        # 進捗率: 0〜1（0%〜100%）
        '<dataValidation type="decimal" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
        'promptTitle="進捗率" prompt="0〜1の値を入力（0.5 = 50%）" '
        'errorTitle="入力エラー" error="進捗率は0〜1の範囲で入力してください（例: 0.5 = 50%）" sqref="G5:G104">'
        "<formula1>0</formula1><formula2>1</formula2>"
        "</dataValidation>"
        # ステータス: 数式で自動計算されるが、バリデーションでリスト制約
        '<dataValidation type="list" allowBlank="1" showDropDown="1" sqref="H5:H104">'
        "<formula1>Config!$F$4:$F$20</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )


def get_template_buttons() -> List[ButtonDefinition]:
    """Template / PRJ シート用のボタン定義を返す。"""
    return [
        ButtonDefinition(
            name="Up",
            macro_name="modWbsCommands.MoveTaskRowUp",
            row=2,  # 3行目 (0-indexed)
            col=0,  # A列
            width=60,
            height=22,
            text="▲ Up",
        ),
        ButtonDefinition(
            name="Down",
            macro_name="modWbsCommands.MoveTaskRowDown",
            row=2,  # 3行目 (0-indexed)
            col=1,  # B列
            width=60,
            height=22,
            text="▼ Down",
        ),
    ]


def get_wbs_column_defs() -> List[ColumnDef]:
    """WBS/Template シート用の列幅定義を返す。"""
    return [
        ColumnDef(min_col=1, max_col=1, width=4),      # A: Lv
        ColumnDef(min_col=2, max_col=2, width=28),     # B: タスク名
        ColumnDef(min_col=3, max_col=3, width=10),     # C: 担当
        ColumnDef(min_col=4, max_col=4, width=11),     # D: 開始日
        ColumnDef(min_col=5, max_col=5, width=6),      # E: 工数
        ColumnDef(min_col=6, max_col=6, width=11),     # F: 終了日
        ColumnDef(min_col=7, max_col=7, width=6),      # G: 進捗率
        ColumnDef(min_col=8, max_col=8, width=9),      # H: ステータス
        ColumnDef(min_col=9, max_col=9, width=15),     # I: 備考
        ColumnDef(min_col=10, max_col=10, width=8),    # J: 全体進捗
        ColumnDef(min_col=11, max_col=70, width=2.5),  # K-BR: ガントチャート日付列（60列）
    ]


def get_config_column_defs() -> List[ColumnDef]:
    """Config シート用の列幅定義を返す。"""
    return [
        ColumnDef(min_col=1, max_col=1, width=12),     # A: ラベル
        ColumnDef(min_col=2, max_col=2, width=15),     # B: 祝日
        ColumnDef(min_col=3, max_col=3, width=3),      # C: 空白
        ColumnDef(min_col=4, max_col=4, width=15),     # D: 担当者
        ColumnDef(min_col=5, max_col=5, width=3),      # E: 空白
        ColumnDef(min_col=6, max_col=6, width=12),     # F: ステータス
    ]


def get_case_master_column_defs() -> List[ColumnDef]:
    """Case_Master シート用の列幅定義を返す。"""
    return [
        ColumnDef(min_col=1, max_col=1, width=12),     # A: 案件ID
        ColumnDef(min_col=2, max_col=2, width=25),     # B: 案件名
        ColumnDef(min_col=3, max_col=3, width=20),     # C: メモ
        ColumnDef(min_col=4, max_col=4, width=10),     # D: 施策数
        ColumnDef(min_col=5, max_col=5, width=10),     # E: 平均進捗
        ColumnDef(min_col=6, max_col=6, width=3),      # F: 空白
        ColumnDef(min_col=7, max_col=7, width=12),     # G: 施策ID
        ColumnDef(min_col=8, max_col=8, width=12),     # H: 親案件ID
        ColumnDef(min_col=9, max_col=9, width=20),     # I: 施策名
        ColumnDef(min_col=10, max_col=10, width=12),   # J: 開始日
        ColumnDef(min_col=11, max_col=11, width=12),   # K: WBSリンク
        ColumnDef(min_col=12, max_col=12, width=12),   # L: WBSシート名
        ColumnDef(min_col=13, max_col=13, width=10),   # M: 実進捗
        ColumnDef(min_col=14, max_col=14, width=20),   # N: 備考
    ]


def get_measure_master_column_defs() -> List[ColumnDef]:
    """Measure_Master シート用の列幅定義を返す。"""
    return [
        ColumnDef(min_col=1, max_col=1, width=12),     # A: 施策ID
        ColumnDef(min_col=2, max_col=2, width=12),     # B: 親案件ID
        ColumnDef(min_col=3, max_col=3, width=25),     # C: 施策名
        ColumnDef(min_col=4, max_col=4, width=12),     # D: 開始日
        ColumnDef(min_col=5, max_col=5, width=12),     # E: WBSリンク
        ColumnDef(min_col=6, max_col=6, width=12),     # F: WBSシート名
        ColumnDef(min_col=7, max_col=7, width=10),     # G: 実進捗
        ColumnDef(min_col=8, max_col=8, width=20),     # H: 備考
    ]


def get_kanban_column_defs() -> List[ColumnDef]:
    """Kanban_View シート用の列幅定義を返す。"""
    return [
        ColumnDef(min_col=1, max_col=1, width=12),     # A: ラベル
        ColumnDef(min_col=2, max_col=2, width=25),     # B: To Do
        ColumnDef(min_col=3, max_col=3, width=3),      # C: 空白
        ColumnDef(min_col=4, max_col=4, width=25),     # D: Doing
        ColumnDef(min_col=5, max_col=5, width=3),      # E: 空白
        ColumnDef(min_col=6, max_col=6, width=25),     # F: Done
    ]


def template_sheet(
    sample: bool = False,
    password_hash: str = "",
    include_buttons: bool = True,
    vml_rid: str | None = None,
) -> str:
    """Template / PRJ シートを生成する。

    編集可能: Lv(A), タスク名(B), 担当(C), 開始日(D), 工数(E), 進捗率(G), 備考(I)
             タスク行 5〜104 行目。行挿入許可。
    保護: 終了日(F), ステータス(H), 全体進捗(J2), ヘッダー(4行目), ガント領域

    Args:
        sample: サンプルデータを含めるか
        password_hash: シート保護用パスワードハッシュ
        include_buttons: Up/Down ボタンを含めるか
        vml_rid: VML描画への参照ID（ボタンを含める場合に指定）
    """
    # セルデータとスタイルマッピングを取得
    cells, cell_styles = template_cells(sample)

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=True)
    legacy_rid = vml_rid if include_buttons else None
    return worksheet_xml(
        cells,
        data_validations=template_data_validations(),
        conditional_formattings=template_conditional_formattings(),
        sheet_protection=protection,
        legacy_drawing_rid=legacy_rid,
        column_defs=get_wbs_column_defs(),
        freeze_row=4,  # ヘッダー行（4行目）まで固定
        freeze_col=2,  # B列まで固定（タスク名を常に表示）
        cell_styles=cell_styles,
    )


def template_conditional_formattings() -> List[str]:
    """条件付き書式の XML を生成する。

    注意: 数式内の < > はXMLエスケープが必要（&lt; &gt;）
    """
    start_row = 5
    end_row = 30  # サンプルデータ分 + 余裕
    gantt_start_col = 11
    gantt_cols = 45  # K-BC（1.5ヶ月分）
    gantt_range = f"{cell_ref(start_row, gantt_start_col)}:{cell_ref(end_row, gantt_start_col + gantt_cols - 1)}"
    col = col_letter(gantt_start_col)

    # XML エスケープ: <> → &lt;&gt;, >= → &gt;=, <= → &lt;=
    gantt_rules = (
        f'<conditionalFormatting sqref="{gantt_range}">'
        f'<cfRule type="expression" dxfId="0" priority="1"><formula>{col}$3=TODAY()</formula></cfRule>'
        f'<cfRule type="expression" dxfId="1" priority="2"><formula>AND($D{start_row}&lt;&gt;"",$E{start_row}&lt;&gt;"",{col}$3&gt;=$D{start_row},{col}$3&lt;=$F{start_row},$H{start_row}="完了")</formula></cfRule>'
        f'<cfRule type="expression" dxfId="2" priority="3"><formula>AND($D{start_row}&lt;&gt;"",$E{start_row}&lt;&gt;"",{col}$3&gt;=$D{start_row},{col}$3&lt;=$F{start_row},$H{start_row}="遅延")</formula></cfRule>'
        f'<cfRule type="expression" dxfId="3" priority="4"><formula>AND($D{start_row}&lt;&gt;"",$E{start_row}&lt;&gt;"",{col}$3&gt;=$D{start_row},{col}$3&lt;=$F{start_row},$H{start_row}&lt;&gt;"",$H{start_row}&lt;&gt;"完了",$H{start_row}&lt;&gt;"遅延")</formula></cfRule>'
        "</conditionalFormatting>"
    )

    status_range = f"H{start_row}:H{end_row}"
    status_rules = (
        f'<conditionalFormatting sqref="{status_range}">'
        f'<cfRule type="expression" dxfId="4" priority="5"><formula>$H{start_row}="未着手"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="5" priority="6"><formula>$H{start_row}="進行中"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="6" priority="7"><formula>$H{start_row}="遅延"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="7" priority="8"><formula>$H{start_row}="完了"</formula></cfRule>'
        "</conditionalFormatting>"
    )

    # Lv1行の強調表示（A列=1の場合、濃紺背景・白太字）
    lv1_range = f"A{start_row}:I{end_row}"
    lv1_rules = (
        f'<conditionalFormatting sqref="{lv1_range}">'
        f'<cfRule type="expression" dxfId="8" priority="9"><formula>$A{start_row}=1</formula></cfRule>'
        "</conditionalFormatting>"
    )

    return [gantt_rules, status_rules, lv1_rules]


def case_master_sheet(password_hash: str = "", m365_mode: bool = False) -> str:
    """Case_Master シートを生成する。

    編集可能: 案件ID(A), 案件名(B), メモ(C) の 2〜100 行目、案件選択(H1)
    保護: 施策数(D), 平均進捗(E), ドリルダウン領域(G3:N104)

    Args:
        password_hash: シート保護用パスワードハッシュ
        m365_mode: True の場合、FILTER を使ったドリルダウン表示
    """
    cells: List[Tuple[int, int, object]] = []
    styles: dict[Tuple[int, int], int] = {}

    # ヘッダー行
    headers = ["案件ID", "案件名", "メモ", "施策数", "平均進捗"]
    for col, header in enumerate(headers, start=1):
        cells.append((1, col, header))
        styles[(1, col)] = STYLE_HEADER

    # 案件データ
    for idx, (case_id, name) in enumerate(CASES, start=2):
        cells.extend([
            (idx, 1, case_id),
            (idx, 2, name),
            (idx, 4, Formula(f"COUNTIF(Measure_Master!$B:$B,A{idx})")),
            (idx, 5, Formula(f"IFERROR(AVERAGEIF(Measure_Master!$B:$B,A{idx},Measure_Master!$G:$G),0)")),
        ])
        # 入力セル
        styles[(idx, 1)] = STYLE_INPUT
        styles[(idx, 2)] = STYLE_INPUT
        styles[(idx, 3)] = STYLE_INPUT
        # 計算セル
        styles[(idx, 4)] = STYLE_CALC
        styles[(idx, 5)] = STYLE_CALC

    # 空の入力行
    for row in range(2 + len(CASES), 12):
        styles[(row, 1)] = STYLE_INPUT
        styles[(row, 2)] = STYLE_INPUT
        styles[(row, 3)] = STYLE_INPUT

    # ドリルダウンエリア
    title = "📋 案件ドリルダウン" + (" (M365版)" if m365_mode else "")
    cells.append((1, 7, title))
    styles[(1, 7)] = STYLE_SUBHEADER
    cells.append((1, 8, "CASE-001"))
    styles[(1, 8)] = STYLE_INPUT

    drill_down_headers = ["施策ID", "親案件ID", "施策名", "開始日", "WBSリンク", "シート名", "実進捗"]
    for col, header in enumerate(drill_down_headers, start=7):
        cells.append((2, col, header))
        styles[(2, col)] = STYLE_HEADER

    if m365_mode:
        # M365版: FILTER を使ったドリルダウン表示（スピル）
        # 選択した案件に紐づく施策をすべて表示
        drilldown_formula = (
            'IF($H$1="","← 案件IDを選択",'
            'IFERROR(FILTER(Measure_Master!$A$2:$H$104,Measure_Master!$B$2:$B$104=$H$1,"該当なし"),""))'
        )
        cells.append((3, 7, Formula(drilldown_formula)))
        styles[(3, 7)] = STYLE_CALC

        # 操作ガイド
        cells.append((4, 7, "💡 H1で案件IDを選択すると施策一覧がスピル表示されます"))
        styles[(4, 7)] = STYLE_DESCRIPTION
    else:
        # 通常版: COUNTIF + 説明表示
        drilldown_formula = (
            'IF($H$1="","← 案件IDを選択",COUNTIF(Measure_Master!$B:$B,$H$1)&" 件の施策")'
        )
        cells.append((3, 7, Formula(drilldown_formula)))
        styles[(3, 7)] = STYLE_CALC

        # 補足説明
        cells.append((3, 8, "→ Measure_Masterで詳細確認"))
        styles[(3, 8)] = STYLE_DESCRIPTION

        # 操作ガイド
        cells.append((4, 7, "💡 H1で案件IDを選択すると施策数を表示"))
        styles[(4, 7)] = STYLE_DESCRIPTION

    data_validations = (
        '<dataValidations count="1">'
        '<dataValidation type="list" allowBlank="1" showDropDown="1" showErrorMessage="1" showInputMessage="1" errorStyle="stop" errorTitle="入力エラー" error="リストから選択してください" promptTitle="案件IDの選択" prompt="プルダウンから案件IDを選択してください" sqref="H1">'
        "<formula1>CaseIds</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=False)
    return worksheet_xml(
        cells,
        data_validations=data_validations,
        sheet_protection=protection,
        column_defs=get_case_master_column_defs(),
        freeze_row=1,
        cell_styles=styles,
    )


def measure_master_sheet(password_hash: str = "") -> str:
    """Measure_Master シートを生成する。

    編集可能: 施策ID(A), 親案件ID(B), 施策名(C), 開始日(D), WBSシート名(F), 備考(H) の 2〜104 行目
    保護: WBSリンク(E), 実進捗(G), ヘッダー行
    """
    cells: List[Tuple[int, int, object]] = []
    styles: dict[Tuple[int, int], int] = {}

    # ヘッダー行
    headers = ["施策ID", "親案件ID", "施策名", "開始日", "WBSリンク", "WBSシート名", "実進捗", "備考"]
    for col, header in enumerate(headers, start=1):
        cells.append((1, col, header))
        styles[(1, col)] = STYLE_HEADER

    # 施策データ
    for idx, (mid, cid, name, start, sheet_name) in enumerate(MEASURES, start=2):
        cells.extend([
            (idx, 1, mid),
            (idx, 2, cid),
            (idx, 3, name),
            (idx, 4, start),
            (idx, 6, sheet_name),
            (idx, 5, Formula(f"HYPERLINK(\"#'\" & F{idx} & \"'!A1\", \"WBSを開く\")")),
            (idx, 7, Formula(f"IF(F{idx}=\"\",\"\",IFERROR(INDIRECT(\"'\" & F{idx} & \"'!J2\"),\"未リンク\"))")),
        ])
        # 入力セル
        for col in [1, 2, 3, 4, 6, 8]:  # A,B,C,D,F,H
            styles[(idx, col)] = STYLE_INPUT
        # 計算/リンクセル
        styles[(idx, 5)] = STYLE_CALC
        styles[(idx, 7)] = STYLE_CALC

    # 空の入力行
    for row in range(2 + len(MEASURES), 12):
        for col in [1, 2, 3, 4, 6, 8]:
            styles[(row, col)] = STYLE_INPUT

    data_validations = (
        '<dataValidations count="1">'
        '<dataValidation type="list" allowBlank="0" showDropDown="1" showErrorMessage="1" showInputMessage="1" errorStyle="stop" errorTitle="入力エラー" error="リスト外の値は入力できません" promptTitle="案件IDの選択" prompt="プルダウンから案件IDを選択してください" sqref="B2:B104">'
        "<formula1>CaseIds</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    # 警告用条件付き書式
    # G列: 「未リンク」時にオレンジ背景（dxfId 9）
    unlinked_warning = (
        '<conditionalFormatting sqref="G2:G104">'
        '<cfRule type="containsText" dxfId="9" priority="1" operator="containsText" text="未リンク">'
        '<formula>NOT(ISERROR(SEARCH("未リンク",G2)))</formula>'
        '</cfRule>'
        '</conditionalFormatting>'
    )
    # B列: 無効な案件ID時に黄色背景（dxfId 10）- 空でなく、Case_Masterに存在しない場合
    invalid_case_warning = (
        '<conditionalFormatting sqref="B2:B104">'
        '<cfRule type="expression" dxfId="10" priority="2">'
        '<formula>AND(B2&lt;&gt;"",ISNA(MATCH(B2,CaseIds,0)))</formula>'
        '</cfRule>'
        '</conditionalFormatting>'
    )

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=False)
    return worksheet_xml(
        cells,
        data_validations=data_validations,
        conditional_formattings=[unlinked_warning, invalid_case_warning],
        sheet_protection=protection,
        column_defs=get_measure_master_column_defs(),
        freeze_row=1,
        cell_styles=styles,
    )


def kanban_sheet(password_hash: str = "", m365_mode: bool = False) -> str:
    """Kanban_View シートを生成する。

    編集可能: B2 (WBS シート名選択) のみ
    保護: カード生成式 (B5:G104)、ヘッダー (1〜4 行)

    Args:
        password_hash: シート保護用パスワードハッシュ
        m365_mode: True の場合、FILTER/LET/MAP を使った詳細カード表示

    カード表示形式:
    - 通常版: 件数のみ
    - M365版: タスク名 + 担当者 + 期限（スピル表示）
    """
    cells: List[Tuple[int, int, object]] = []
    styles: dict[Tuple[int, int], int] = {}

    # タイトル
    title = "📋 カンバンビュー" + (" (M365版)" if m365_mode else "")
    cells.append((1, 1, title))
    styles[(1, 1)] = STYLE_TITLE

    # 施策選択
    cells.append((2, 1, "施策を選択:"))
    styles[(2, 1)] = STYLE_SUBHEADER
    cells.append((2, 2, "PRJ_001"))
    styles[(2, 2)] = STYLE_INPUT

    # 操作ガイド
    guide_text = "💡 B2で WBS シートを選択 → タスクカード表示。ダブルクリックでステータス変更。"
    if m365_mode:
        guide_text = "💡 B2で WBS シート選択 → タスクが「名前/担当/期限」形式でスピル表示。"
    cells.append((3, 1, guide_text))
    styles[(3, 1)] = STYLE_DESCRIPTION

    # カンバンヘッダー
    cells.append((4, 2, "📥 To Do"))
    styles[(4, 2)] = STYLE_HEADER
    cells.append((4, 4, "🔄 Doing"))
    styles[(4, 4)] = STYLE_HEADER
    cells.append((4, 6, "✅ Done"))
    styles[(4, 6)] = STYLE_HEADER
    cells.append((4, 8, "⚠️ 遅延"))
    styles[(4, 8)] = STYLE_HEADER

    if m365_mode:
        # M365版: FILTER/LET を使った詳細カード表示
        # タスク名 + 担当者 + 期限 + 進捗率 をカード形式で表示
        # アイコン付きで視認性向上

        # カード形式:
        # タスク名
        # 👤 担当者
        # 📅 期限: yyyy/mm/dd
        # 📊 進捗: XX%

        # To Do: 未着手タスク
        todo_formula = (
            'IFERROR(LET('
            '_sheet,$B$2,'
            '_tasks,INDIRECT("\'"&_sheet&"\'!B5:B104"),'
            '_owners,INDIRECT("\'"&_sheet&"\'!C5:C104"),'
            '_ends,INDIRECT("\'"&_sheet&"\'!F5:F104"),'
            '_progress,INDIRECT("\'"&_sheet&"\'!G5:G104"),'
            '_status,INDIRECT("\'"&_sheet&"\'!H5:H104"),'
            '_card,_tasks&CHAR(10)&"👤 "&_owners&CHAR(10)&"📅 "&TEXT(_ends,"m/d")&" | 📊 "&TEXT(_progress,"0%"),'
            '_filtered,FILTER(_card,(_status="未着手")*(_tasks<>""),""),'
            '_filtered'
            '),"")'
        )
        cells.append((5, 2, Formula(todo_formula)))
        styles[(5, 2)] = STYLE_CALC

        # Doing: 進行中タスク
        doing_formula = (
            'IFERROR(LET('
            '_sheet,$B$2,'
            '_tasks,INDIRECT("\'"&_sheet&"\'!B5:B104"),'
            '_owners,INDIRECT("\'"&_sheet&"\'!C5:C104"),'
            '_ends,INDIRECT("\'"&_sheet&"\'!F5:F104"),'
            '_progress,INDIRECT("\'"&_sheet&"\'!G5:G104"),'
            '_status,INDIRECT("\'"&_sheet&"\'!H5:H104"),'
            '_card,_tasks&CHAR(10)&"👤 "&_owners&CHAR(10)&"📅 "&TEXT(_ends,"m/d")&" | 📊 "&TEXT(_progress,"0%"),'
            '_filtered,FILTER(_card,(_status="進行中")*(_tasks<>""),""),'
            '_filtered'
            '),"")'
        )
        cells.append((5, 4, Formula(doing_formula)))
        styles[(5, 4)] = STYLE_CALC

        # Done: 完了タスク
        done_formula = (
            'IFERROR(LET('
            '_sheet,$B$2,'
            '_tasks,INDIRECT("\'"&_sheet&"\'!B5:B104"),'
            '_owners,INDIRECT("\'"&_sheet&"\'!C5:C104"),'
            '_ends,INDIRECT("\'"&_sheet&"\'!F5:F104"),'
            '_progress,INDIRECT("\'"&_sheet&"\'!G5:G104"),'
            '_status,INDIRECT("\'"&_sheet&"\'!H5:H104"),'
            '_card,_tasks&CHAR(10)&"👤 "&_owners&CHAR(10)&"📅 "&TEXT(_ends,"m/d")&" | 📊 "&TEXT(_progress,"0%"),'
            '_filtered,FILTER(_card,(_status="完了")*(_tasks<>""),""),'
            '_filtered'
            '),"")'
        )
        cells.append((5, 6, Formula(done_formula)))
        styles[(5, 6)] = STYLE_CALC

        # 遅延タスク
        delay_formula = (
            'IFERROR(LET('
            '_sheet,$B$2,'
            '_tasks,INDIRECT("\'"&_sheet&"\'!B5:B104"),'
            '_owners,INDIRECT("\'"&_sheet&"\'!C5:C104"),'
            '_ends,INDIRECT("\'"&_sheet&"\'!F5:F104"),'
            '_progress,INDIRECT("\'"&_sheet&"\'!G5:G104"),'
            '_status,INDIRECT("\'"&_sheet&"\'!H5:H104"),'
            '_card,_tasks&CHAR(10)&"👤 "&_owners&CHAR(10)&"📅 "&TEXT(_ends,"m/d")&" | 📊 "&TEXT(_progress,"0%"),'
            '_filtered,FILTER(_card,(_status="遅延")*(_tasks<>""),""),'
            '_filtered'
            '),"")'
        )
        cells.append((5, 8, Formula(delay_formula)))
        styles[(5, 8)] = STYLE_CALC
    else:
        # 通常版: シンプルなCOUNTIF + 件数表示
        # To Do: 未着手タスク件数
        todo_formula = (
            'IF($B$2="","",IFERROR(COUNTIF(INDIRECT("\'"&$B$2&"\'!H5:H104"),"未着手")&" 件",""))'
        )
        cells.append((5, 2, Formula(todo_formula)))
        styles[(5, 2)] = STYLE_CALC

        # Doing: 進行中タスク件数
        doing_formula = (
            'IF($B$2="","",IFERROR(COUNTIF(INDIRECT("\'"&$B$2&"\'!H5:H104"),"進行中")&" 件",""))'
        )
        cells.append((5, 4, Formula(doing_formula)))
        styles[(5, 4)] = STYLE_CALC

        # Done: 完了タスク件数
        done_formula = (
            'IF($B$2="","",IFERROR(COUNTIF(INDIRECT("\'"&$B$2&"\'!H5:H104"),"完了")&" 件",""))'
        )
        cells.append((5, 6, Formula(done_formula)))
        styles[(5, 6)] = STYLE_CALC

        # 遅延タスク件数
        delay_formula = (
            'IF($B$2="","",IFERROR(COUNTIF(INDIRECT("\'"&$B$2&"\'!H5:H104"),"遅延")&" 件",""))'
        )
        cells.append((5, 8, Formula(delay_formula)))
        styles[(5, 8)] = STYLE_CALC

    data_validations = (
        '<dataValidations count="1">'
        '<dataValidation type="list" allowBlank="1" showDropDown="1" showErrorMessage="1" showInputMessage="1" errorStyle="stop" errorTitle="入力エラー" error="リスト外の値は入力できません" promptTitle="WBS シート名の選択" prompt="プルダウンから施策の WBS シート名を選択してください" sqref="B2">'
        "<formula1>Measure_Master!$F$2:$F$20</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=False)
    return worksheet_xml(
        cells,
        data_validations=data_validations,
        sheet_protection=protection,
        column_defs=get_kanban_column_defs(),
        freeze_row=4,
        cell_styles=styles,
    )


# --------------------------- レポート生成 ---------------------------

def generate_report_lines(
    project_count: int,
    sample_first_project: bool,
    sample_all_projects: bool,
    workbook_path: Path,
) -> List[str]:
    """ブック構成と進捗状況を日本語でまとめたレポートを返す。"""

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    has_sample = sample_first_project or sample_all_projects

    lines = [
        "=" * 50,
        "Modern Excel PMS 生成レポート",
        "=" * 50,
        "",
        "## 基本情報",
        f"生成日時: {generated_at}",
        f"ブック出力先: {workbook_path}",
        f"PRJ シート数: {project_count}",
        f"サンプルデータ: {'全てのPRJに配置' if sample_all_projects else ('最初の1枚に配置' if sample_first_project else 'なし')}",
    ]

    # サンプルデータがある場合は進捗分析を追加
    if has_sample:
        lines.append("")
        lines.append("-" * 50)
        lines.append("## 進捗サマリー (サンプルデータ)")
        lines.append("-" * 50)

        # 全体進捗率（工数加重平均）
        overall_progress = calculate_weighted_progress(SAMPLE_TASKS)
        total_effort = sum(t.effort for t in SAMPLE_TASKS)
        completed_effort = sum(t.effort * t.progress for t in SAMPLE_TASKS)

        lines.append("")
        lines.append(f"全体進捗率: {overall_progress:.1%}")
        lines.append(f"  - 総工数: {total_effort} 人日")
        lines.append(f"  - 消化工数: {completed_effort:.1f} 人日")

        # ステータス別集計
        status_counts = count_by_status(SAMPLE_TASKS)
        total_tasks = len(SAMPLE_TASKS)
        completed_tasks = status_counts.get("完了", 0)

        lines.append("")
        lines.append("ステータス別タスク数:")
        for status in STATUSES:
            count = status_counts.get(status, 0)
            pct = count / total_tasks * 100 if total_tasks > 0 else 0
            bar = "#" * int(pct / 5)  # 5% ごとに # 1個
            lines.append(f"  {status:6s}: {count:2d} ({pct:5.1f}%) {bar}")

        # 案件消化度
        lines.append("")
        lines.append(f"タスク完了率: {completed_tasks}/{total_tasks} ({completed_tasks/total_tasks:.1%})")

        # 施策別進捗（PRJ_001 のみサンプルがある想定）
        lines.append("")
        lines.append("施策別進捗:")
        for mid, cid, name, start, sheet_name in MEASURES:
            if sheet_name == "PRJ_001":
                prj_progress = overall_progress
                lines.append(f"  - {mid} ({name}): {prj_progress:.1%}")
            else:
                lines.append(f"  - {mid} ({name}): -- (データなし)")

        # 担当者別負荷
        owner_effort: dict[str, int] = {}
        owner_completed: dict[str, float] = {}
        for task in SAMPLE_TASKS:
            owner_effort[task.owner] = owner_effort.get(task.owner, 0) + task.effort
            owner_completed[task.owner] = owner_completed.get(task.owner, 0) + task.effort * task.progress

        lines.append("")
        lines.append("担当者別負荷:")
        for owner in sorted(owner_effort.keys()):
            effort = owner_effort[owner]
            completed = owner_completed[owner]
            pct = completed / effort if effort > 0 else 0
            lines.append(f"  - {owner}: {effort} 人日 (消化 {pct:.1%})")

    lines.append("")
    lines.append("-" * 50)
    lines.append("## マスターデータ")
    lines.append("-" * 50)

    lines.append("")
    lines.append("案件一覧:")
    for case_id, name in CASES:
        # 案件に紐づく施策数を計算
        measure_count = sum(1 for m in MEASURES if m[1] == case_id)
        lines.append(f"  - {case_id}: {name} (施策数: {measure_count})")

    lines.append("")
    lines.append("施策一覧:")
    for mid, cid, name, start, sheet_name in MEASURES:
        lines.append(f"  - {mid} ({cid}) {name}")
        lines.append(f"      開始日: {start} / WBS: {sheet_name}")

    lines.append("")
    lines.append("ステータス候補:")
    for status in STATUSES:
        lines.append(f"  - {status}")

    lines.append("")
    lines.append("担当者マスタ:")
    for member in MEMBERS:
        lines.append(f"  - {member}")

    lines.append("")
    lines.append("=" * 50)

    return lines


def write_report_text(lines: Sequence[str], output_path: Path) -> None:
    """レポートテキストを UTF-8 で書き出す。"""

    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _escape_pdf_text(text: str) -> str:
    """PDF 文字列リテラル向けのエスケープ処理。"""

    sanitized = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    return sanitized


def export_pdf_report(lines: Sequence[str], output_path: Path) -> None:
    """標準フォントのみで構成したシンプルな PDF を生成する。"""

    page_height = 842  # A4 高さ (pt)
    margin_left = 50
    margin_top = 50
    line_height = 16

    content_lines: List[str] = ["BT", "/F1 12 Tf"]
    y_cursor = page_height - margin_top
    for line in lines:
        escaped = _escape_pdf_text(line)
        content_lines.append(f"1 0 0 1 {margin_left} {y_cursor} Tm ({escaped}) Tj")
        y_cursor -= line_height
        if y_cursor < margin_top:
            break  # 1 ページのみサポート
    content_lines.append("ET")
    content_stream = "\n".join(content_lines).encode("utf-8")

    objects: List[bytes] = []
    objects.append(b"1 0 obj<< /Type /Catalog /Pages 2 0 R >>endobj\n")
    objects.append(b"2 0 obj<< /Type /Pages /Count 1 /Kids [3 0 R] >>endobj\n")
    objects.append(
        b"3 0 obj<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
        b"/Contents 4 0 R /Resources<< /Font << /F1 5 0 R >> >> >>endobj\n"
    )
    objects.append(
        f"4 0 obj<< /Length {len(content_stream)} >>stream\n".encode("utf-8")
        + content_stream
        + b"\nendstream\nendobj\n"
    )
    objects.append(b"5 0 obj<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>endobj\n")

    # クロスリファレンスを組み立てる
    offsets = []
    position = len(b"%PDF-1.4\n")
    for obj in objects:
        offsets.append(position)
        position += len(obj)

    xref_entries = ["0000000000 65535 f "]
    for offset in offsets:
        xref_entries.append(f"{offset:010} 00000 n ")
    xref_content = "\n".join(xref_entries) + "\n"

    trailer = (
        f"<< /Root 1 0 R /Size {len(objects) + 1} >>\nstartxref\n{position}\n%%EOF"
    )

    pdf_binary = b"".join(
        [
            b"%PDF-1.4\n",
            *objects,
            b"xref\n0 ",
            str(len(objects) + 1).encode("utf-8"),
            b"\n",
            xref_content.encode("utf-8"),
            b"trailer\n",
            trailer.encode("utf-8"),
        ]
    )

    output_path.write_bytes(pdf_binary)


# --------------------------- メイン ---------------------------

def build_workbook(
    project_count: int,
    sample_first_project: bool,
    sample_all_projects: bool,
    output_path: Path,
    include_vba: bool = False,
    include_buttons: bool = False,
    regenerate_vba: bool = False,
    m365_mode: bool = False,
) -> List[str]:
    """指定した枚数の PRJ シートを生成してブックを書き出し、レポート用テキストを返す。

    Args:
        project_count: 生成する PRJ シート数
        sample_first_project: 最初の PRJ にサンプルデータを含めるか
        sample_all_projects: 全 PRJ にサンプルデータを含めるか
        output_path: 出力先パス
        include_vba: VBA プロジェクトを含めるか
        include_buttons: Up/Down ボタンを含めるか
        regenerate_vba: vbaProject.bin を強制的に再生成するか
        m365_mode: Microsoft 365 専用機能（FILTER/LET/MAP）を使用するか
    """

    # パスワードハッシュを計算
    password = get_sheet_password()
    pwd_hash = excel_password_hash(password)

    # ボタン定義を取得
    buttons = get_template_buttons() if include_buttons else []
    vml_rid = "rId1"

    # VML ファイルとリレーションシップを追跡
    # key: sheet_index (1-based), value: (vml_filename, sheet_name)
    vml_sheets: dict[int, Tuple[str, str]] = {}

    # Config / Template
    sheet_names = ["Config", "Template"]
    sheets_xml: List[str] = [
        config_sheet(password_hash=pwd_hash),
        template_sheet(sample=False, password_hash=pwd_hash, include_buttons=include_buttons, vml_rid=vml_rid if include_buttons else None),
    ]
    # Template シート (sheet2) にボタンを追加
    if include_buttons:
        vml_sheets[2] = ("vmlDrawing1.vml", "Template")

    # PRJ_xxx をまとめて生成
    vml_index = 2  # vmlDrawing2.vml から開始
    for idx in range(1, project_count + 1):
        sheet_name = f"PRJ_{idx:03d}"
        sheet_names.append(sheet_name)
        is_sample = sample_all_projects or (sample_first_project and idx == 1)
        sheets_xml.append(template_sheet(sample=is_sample, password_hash=pwd_hash, include_buttons=include_buttons, vml_rid=vml_rid if include_buttons else None))
        # PRJ シートにもボタンを追加
        if include_buttons:
            sheet_index = 2 + idx  # Config=1, Template=2, PRJ_001=3, ...
            vml_sheets[sheet_index] = (f"vmlDrawing{vml_index}.vml", sheet_name)
            vml_index += 1

    # 末尾のマスターシート群
    sheet_names.extend(["Case_Master", "Measure_Master", "Kanban_View"])
    sheets_xml.extend([
        case_master_sheet(password_hash=pwd_hash, m365_mode=m365_mode),
        measure_master_sheet(password_hash=pwd_hash),
        kanban_sheet(password_hash=pwd_hash, m365_mode=m365_mode),
    ])

    defined_names = {
        "CaseIds": "Case_Master!$A$2:$A$100",
        "MeasureList": "Measure_Master!$A$2:$H$104",
        "CaseDrilldownArea": "Case_Master!$G$3:$N$104",
    }

    has_vml = len(vml_sheets) > 0

    # VBAバイナリを事前に取得（テンプレートがない場合は自動生成）
    vba_binary: bytes | None = None
    actual_has_vba = False
    if include_vba:
        vba_modules = load_vba_modules()
        vba_binary = vba_project_binary(vba_modules, regenerate=regenerate_vba)
        actual_has_vba = vba_binary is not None

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets_xml), has_vml=has_vml, has_vba=actual_has_vba))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names, defined_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets_xml), has_vba=actual_has_vba))
        zf.writestr("xl/styles.xml", styles_xml())

        if actual_has_vba and vba_binary:
            zf.writestr("xl/vbaProject.bin", vba_binary)

        for idx, xml in enumerate(sheets_xml, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", xml)

            # ボタン付きシートの場合、VML ファイルとリレーションシップを書き込む
            if idx in vml_sheets:
                vml_filename, sheet_name_for_vml = vml_sheets[idx]
                # VML 描画ファイル
                vml_xml = vml_drawing_xml(buttons, sheet_name_for_vml)
                zf.writestr(f"xl/drawings/{vml_filename}", vml_xml)
                # ワークシートリレーションシップ
                rels_xml = worksheet_rels_xml(vml_rid, vml_filename)
                if rels_xml:
                    zf.writestr(f"xl/worksheets/_rels/sheet{idx}.xml.rels", rels_xml)

    ext = output_path.suffix.lower()
    if actual_has_vba:
        file_type = "マクロ有効ブック (.xlsm)"
    elif include_vba:
        file_type = "マクロ有効ブック (.xlsm) - VBAは手動追加が必要"
    else:
        file_type = "通常ブック (.xlsx)"

    m365_note = " [M365専用: FILTER/LET対応]" if m365_mode else ""
    print(f"ブックを生成しました: {output_path} ({file_type}){m365_note}")

    return generate_report_lines(project_count, sample_first_project, sample_all_projects, output_path)


def main() -> None:
    # デフォルト出力パスを .xlsx に変更（VBA なしの場合）
    default_output = Path(__file__).resolve().parent.parent / "ModernExcelPMS.xlsx"

    parser = argparse.ArgumentParser(description="Modern Excel PMS 雛形を生成する")
    parser.add_argument("--projects", type=int, default=2, help="生成する PRJ_xxx シート数")
    parser.add_argument(
        "--sample-first",
        action="store_true",
        default=True,
        help="最初の PRJ シートにサンプルタスクを埋め込む (デフォルト: True)",
    )
    parser.add_argument(
        "--no-sample",
        action="store_true",
        help="サンプルタスクを埋め込まない",
    )
    parser.add_argument(
        "--sample-all",
        action="store_true",
        help="全ての PRJ シートにサンプルタスクを埋め込む",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=default_output,
        help="出力先パス (.xlsx または .xlsm)",
    )
    parser.add_argument(
        "--with-vba",
        action="store_true",
        help="VBA プロジェクトを含める（実験的、Excel で開けない可能性あり）",
    )
    parser.add_argument(
        "--with-buttons",
        action="store_true",
        help="Up/Down ボタン (VML) を含める（実験的）",
    )
    parser.add_argument(
        "--regenerate-vba",
        action="store_true",
        help="vbaProject.bin を強制的に再生成する",
    )
    parser.add_argument(
        "--m365",
        action="store_true",
        default=True,
        help="Microsoft 365 専用版: FILTER/LET/MAP を使用（デフォルト有効）",
    )
    parser.add_argument(
        "--legacy",
        action="store_true",
        help="旧互換モード: FILTER/LET を使用せず COUNTIF で簡略表示（非推奨）",
    )
    parser.add_argument(
        "--report-output",
        type=Path,
        help="ブック構成レポートを書き出すパス (.md や .txt を想定)",
    )
    parser.add_argument(
        "--pdf-output",
        type=Path,
        help="レポート PDF を書き出すパス",
    )
    args = parser.parse_args()

    sample_first = args.sample_first and not args.no_sample
    # --legacy フラグが指定された場合は M365 モードを無効化
    m365_mode = args.m365 and not args.legacy
    report_lines = build_workbook(
        args.projects,
        sample_first,
        args.sample_all,
        args.output,
        include_vba=args.with_vba,
        include_buttons=args.with_buttons,
        regenerate_vba=args.regenerate_vba,
        m365_mode=m365_mode,
    )

    if args.report_output:
        write_report_text(report_lines, args.report_output)
        print(f"レポートを出力しました: {args.report_output}")

    if args.pdf_output:
        export_pdf_report(report_lines, args.pdf_output)
        print(f"PDF レポートを出力しました: {args.pdf_output}")


if __name__ == "__main__":
    main()
