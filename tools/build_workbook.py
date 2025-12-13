"""
Modern Excel PMS の雛形ブックを自動生成するスクリプト。
外部ライブラリに依存せず、OpenXML を直接書き出して `ModernExcelPMS.xlsm` を生成する。
"""

from __future__ import annotations

from dataclasses import dataclass
import argparse
from datetime import datetime
import os
from pathlib import Path
from typing import List, Mapping, Sequence, Tuple
from xml.sax.saxutils import escape
import zipfile

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
MEMBERS = ["PM_佐藤", "TL_田中", "DEV_鈴木", "QA_伊藤"]
STATUSES = ["未着手", "進行中", "遅延", "完了"]
CASES = [("CASE-001", "Web 刷新案件"), ("CASE-002", "新規 SFA 導入")]
MEASURES = [
    ("ME-001", "CASE-001", "LP 改修", "2024-05-07", "PRJ_001"),
    ("ME-002", "CASE-002", "SFA 導入 PoC", "2024-05-13", "PRJ_002"),
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
SAMPLE_TASKS: List[SampleTask] = [
    SampleTask(lv=1, name="キックオフ準備", owner="PM_佐藤", start_date="2024-05-07", effort=2, progress=1.0),
    SampleTask(lv=2, name="要件定義ワークショップ", owner="TL_田中", start_date="2024-05-09", effort=3, progress=0.5),
    SampleTask(lv=2, name="WBS 詳細化", owner="DEV_鈴木", start_date="2024-05-13", effort=5, progress=0.0),
]


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
    """
    ref = cell_ref(row, col)
    style_attr = f' s="{style_id}"' if style_id else ""
    if isinstance(value, Formula):
        return f"<c r=\"{ref}\"{style_attr}><f>{escape(value.expr)}</f></c>"
    if isinstance(value, str):
        return f"<c r=\"{ref}\"{style_attr} t=\"inlineStr\"><is><t>{escape(value)}</t></is></c>"
    if value is None:
        return ""
    return f"<c r=\"{ref}\"{style_attr}><v>{value}</v></c>"


# スタイル ID 定数
STYLE_LOCKED = 0
STYLE_UNLOCKED = 1


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


def worksheet_xml(
    cells: Sequence[Tuple[int, int, object]],
    data_validations: str | None = None,
    conditional_formattings: Sequence[str] | None = None,
    sheet_protection: SheetProtection | None = None,
    unlocked_cells: set[Tuple[int, int]] | None = None,
) -> str:
    """ワークシート XML を生成する。

    Args:
        cells: (row, col, value) のセルデータ
        data_validations: データ検証 XML
        conditional_formattings: 条件付き書式 XML リスト
        sheet_protection: シート保護設定
        unlocked_cells: ロック解除するセルの (row, col) セット
    """
    rows = {}
    for row, col, value in cells:
        rows.setdefault(row, {})[col] = value

    unlocked = unlocked_cells or set()

    xml_lines: List[str] = [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheetData>",
    ]

    for row_idx in sorted(rows):
        xml_lines.append(f"<row r=\"{row_idx}\">")
        for col_idx in sorted(rows[row_idx]):
            style_id = STYLE_UNLOCKED if (row_idx, col_idx) in unlocked else STYLE_LOCKED
            xml_lines.append(cell_xml(row_idx, col_idx, rows[row_idx][col_idx], style_id))
        xml_lines.append("</row>")

    xml_lines.append("</sheetData>")

    if sheet_protection:
        xml_lines.append(sheet_protection.to_xml())

    if data_validations:
        xml_lines.append(data_validations)

    if conditional_formattings:
        xml_lines.extend(conditional_formattings)

    xml_lines.append("</worksheet>")
    return "".join(xml_lines)


def content_types_xml(sheet_count: int) -> str:
    overrides = "".join(
        f"<Override PartName='/xl/worksheets/sheet{idx}.xml' "
        "ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>"
        for idx in range(1, sheet_count + 1)
    )
    return (
        "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>"
        "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>"
        "<Default Extension='xml' ContentType='application/xml'/>"
        "<Override PartName='/xl/workbook.xml' "
        "ContentType='application/vnd.ms-excel.sheet.macroEnabled.main+xml'/>"
        f"{overrides}"
        "<Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/>"
        "<Override PartName='/xl/vbaProject.bin' ContentType='application/vnd.ms-office.vbaProject'/>"
        "</Types>"
    )


def root_rels_xml() -> str:
    return (
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>"
        "</Relationships>"
    )


def workbook_xml(sheet_names: Sequence[str], defined_names: Mapping[str, str] | None = None) -> str:
    sheets_xml = "".join(
        f"<sheet name='{escape(name)}' sheetId='{idx}' r:id='rId{idx}'/>"
        for idx, name in enumerate(sheet_names, start=1)
    )

    defined_names_xml = ""
    if defined_names:
        defined_names_xml = "<definedNames>" + "".join(
            f"<definedName name='{escape(name)}'>{escape(ref)}</definedName>"
            for name, ref in defined_names.items()
        ) + "</definedNames>"

    return (
        "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>"
        f"<sheets>{sheets_xml}</sheets>"
        f"{defined_names_xml}"
        "</workbook>"
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = "".join(
        f"<Relationship Id='rId{idx}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet{idx}.xml'/>"
        for idx in range(1, sheet_count + 1)
    )
    rels += (
        f"<Relationship Id='rId{sheet_count + 1}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>"
    )
    rels += (
        f"<Relationship Id='rId{sheet_count + 2}' Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject' Target='vbaProject.bin'/>"
    )
    return f"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>{rels}</Relationships>"


def styles_xml() -> str:
    """スタイルシートを生成する。

    セルスタイル:
      - xfId=0: 標準（ロック）
      - xfId=1: ロック解除（編集可能セル用）
    """
    return (
        "<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
        "<fonts count='1'><font><sz val='11'/><color theme='1'/><name val='Calibri'/><family val='2'/><scheme val='minor'/></font></fonts>"
        "<fills count='2'><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill></fills>"
        "<borders count='1'><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
        "<cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs>"
        # cellXfs: xfId=0 はロック（デフォルト）、xfId=1 はロック解除
        "<cellXfs count='2'>"
        "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>"
        "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0' applyProtection='1'><protection locked='0'/></xf>"
        "</cellXfs>"
        "<cellStyles count='1'><cellStyle name='標準' xfId='0' builtinId='0'/></cellStyles>"
        "<dxfs count='8'>"
        "<dxf><border><right style='medium'><color rgb='FFE74C3C'/></right></border></dxf>"
        "<dxf><fill><patternFill patternType='solid'><fgColor rgb='FF95A5A6'/><bgColor indexed='64'/></patternFill></fill></dxf>"
        "<dxf><fill><patternFill patternType='solid'><fgColor rgb='FFE74C3C'/><bgColor indexed='64'/></patternFill></fill></dxf>"
        "<dxf><fill><patternFill patternType='solid'><fgColor rgb='FF3498DB'/><bgColor indexed='64'/></patternFill></fill></dxf>"
        "<dxf><fill><patternFill patternType='solid'><fgColor rgb='FFBDC3C7'/><bgColor indexed='64'/></patternFill></fill></dxf>"
        "<dxf><font><color rgb='FFFFFFFF'/></font><fill><patternFill patternType='solid'><fgColor rgb='FF3498DB'/><bgColor indexed='64'/></patternFill></fill></dxf>"
        "<dxf><font><color rgb='FFFFFFFF'/></font><fill><patternFill patternType='solid'><fgColor rgb='FFE74C3C'/><bgColor indexed='64'/></patternFill></fill></dxf>"
        "<dxf><font><color rgb='FFFFFFFF'/></font><fill><patternFill patternType='solid'><fgColor rgb='FF2ECC71'/><bgColor indexed='64'/></patternFill></fill></dxf>"
        "</dxfs>"
        "<tableStyles count='0' defaultTableStyle='TableStyleMedium9' defaultPivotStyle='PivotStyleLight16'/>"
        "</styleSheet>"
    )


def load_vba_modules() -> Mapping[str, str]:
    modules: dict[str, str] = {}
    if not VBA_SOURCE_DIR.exists():
        return modules

    for path in sorted(VBA_SOURCE_DIR.glob("*.bas")):
        modules[path.stem] = path.read_text(encoding="utf-8")

    for path in sorted(VBA_SOURCE_DIR.glob("*.cls")):
        modules[path.stem] = path.read_text(encoding="utf-8")

    return modules


def vba_project_binary(modules: Mapping[str, str]) -> bytes:
    header = ["' Modern Excel PMS VBA project", "' 自動生成される各モジュールの内容"]
    lines: List[str] = header.copy()

    for name, body in modules.items():
        normalized_body = body.replace("\r\n", "\n").replace("\r", "\n")
        lines.append(f"' ----- {name} -----")
        lines.append(normalized_body)
        lines.append("'")

    return "\r\n".join(lines).encode("utf-8")


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
    cells = [
        (1, 1, "祝日リスト"),
        (2, 1, "日付"),
        (1, 4, "担当者マスタ"),
        (2, 4, "氏名"),
        (1, 6, "ステータス候補"),
        (2, 6, "値"),
    ]
    for idx, day in enumerate(HOLIDAYS, start=4):
        cells.append((idx, 2, day))
    for idx, member in enumerate(MEMBERS, start=4):
        cells.append((idx, 4, member))
    for idx, status in enumerate(STATUSES, start=4):
        cells.append((idx, 6, status))

    # ロック解除セル: B4:B200, D4:D200, F4:F200
    unlocked: set[Tuple[int, int]] = set()
    for row in range(4, 201):
        unlocked.add((row, 2))  # B 列
        unlocked.add((row, 4))  # D 列
        unlocked.add((row, 6))  # F 列

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=True)
    return worksheet_xml(cells, sheet_protection=protection, unlocked_cells=unlocked)


def template_cells(sample: bool = False) -> List[Tuple[int, int, object]]:
    cells: List[Tuple[int, int, object]] = [(1, 1, "プロジェクト名"), (1, 2, "Modern Excel PMS")]
    headers = ["Lv", "タスク名", "担当", "開始日", "工数(日)", "終了日", "進捗率", "ステータス", "備考"]
    for col, header in enumerate(headers, start=1):
        cells.append((4, col, header))

    cells.append((1, 11, "ガント開始日"))
    cells.append((2, 11, Formula("TODAY()-3")))

    gantt_start_col = 11
    gantt_columns = 30
    for offset in range(gantt_columns):
        cells.append((3, gantt_start_col + offset, Formula(f"IF($K$2=\"\",\"\",$K$2+{offset})")))

    for row in range(5, 14):
        cells.append((row, 6, Formula(f"IF(OR(D{row}='',E{row}=''),' ',WORKDAY(D{row},E{row}-1,Config!$B$4:$B$20))")))
        cells.append((row, 8, Formula(f"IFS(G{row}=1,'完了',AND(F{row}<TODAY(),G{row}<1),'遅延',AND(D{row}<=TODAY(),G{row}<1),'進行中',TRUE,'未着手')")))

    cells.append((1, 10, "全体進捗"))
    cells.append(
        (
            2,
            10,
            Formula(
                "LET(_eff,E5:E104,_prg,G5:G104,_total,SUM(_eff),IF(_total=0,0,SUMPRODUCT(_eff,_prg)/_total))"
            ),
        )
    )

    if sample:
        for row_offset, task in enumerate(SAMPLE_TASKS):
            row = 5 + row_offset
            cells.extend(
                [
                    (row, 1, task.lv),
                    (row, 2, task.name),
                    (row, 3, task.owner),
                    (row, 4, task.start_date),
                    (row, 5, task.effort),
                    (row, 7, task.progress),
                ]
            )
    return cells


def template_data_validations() -> str:
    return (
        "<dataValidations count='2'>"
        "<dataValidation type='list' allowBlank='1' showDropDown='1' sqref='C5:C104'>"
        "<formula1>Config!$D$4:$D$20</formula1>"
        "</dataValidation>"
        "<dataValidation type='list' allowBlank='1' showDropDown='1' sqref='H5:H104'>"
        "<formula1>Config!$F$4:$F$20</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )


def template_sheet(sample: bool = False, password_hash: str = "") -> str:
    """Template / PRJ シートを生成する。

    編集可能: Lv(A), タスク名(B), 担当(C), 開始日(D), 工数(E), 進捗率(G), ステータス(H), 備考(I)
             タスク行 5〜104 行目。行挿入許可。
    保護: 終了日(F), 全体進捗(J2), ヘッダー(4行目), ガント領域
    """
    # ロック解除セル: A,B,C,D,E,G,H,I 列の 5〜104 行目
    unlocked: set[Tuple[int, int]] = set()
    editable_cols = [1, 2, 3, 4, 5, 7, 8, 9]  # A,B,C,D,E,G,H,I
    for row in range(5, 105):
        for col in editable_cols:
            unlocked.add((row, col))

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=True)
    return worksheet_xml(
        template_cells(sample),
        data_validations=template_data_validations(),
        conditional_formattings=template_conditional_formattings(),
        sheet_protection=protection,
        unlocked_cells=unlocked,
    )


def template_conditional_formattings() -> List[str]:
    start_row = 5
    end_row = 104
    gantt_start_col = 11
    gantt_cols = 30
    gantt_range = f"{cell_ref(start_row, gantt_start_col)}:{cell_ref(end_row, gantt_start_col + gantt_cols - 1)}"
    start_col_letter = col_letter(gantt_start_col)

    gantt_rules = f"""
<conditionalFormatting sqref='{gantt_range}'>
<cfRule type='expression' dxfId='0' priority='1'><formula>{start_col_letter}$3=TODAY()</formula></cfRule>
<cfRule type='expression' dxfId='1' priority='2'><formula>AND($D{start_row}<>"",$E{start_row}<>"",{start_col_letter}$3>=$D{start_row},{start_col_letter}$3<=$F{start_row},$H{start_row}="完了")</formula></cfRule>
<cfRule type='expression' dxfId='2' priority='3'><formula>AND($D{start_row}<>"",$E{start_row}<>"",{start_col_letter}$3>=$D{start_row},{start_col_letter}$3<=$F{start_row},$H{start_row}="遅延")</formula></cfRule>
<cfRule type='expression' dxfId='3' priority='4'><formula>AND($D{start_row}<>"",$E{start_row}<>"",{start_col_letter}$3>=$D{start_row},{start_col_letter}$3<=$F{start_row},$H{start_row}<>"",$H{start_row}<>"完了",$H{start_row}<>"遅延")</formula></cfRule>
</conditionalFormatting>
"""

    status_range = f"{cell_ref(start_row, 8)}:{cell_ref(end_row, 8)}"
    status_rules = f"""
<conditionalFormatting sqref='{status_range}'>
<cfRule type='expression' dxfId='4' priority='5'><formula>$H{start_row}="未着手"</formula></cfRule>
<cfRule type='expression' dxfId='5' priority='6'><formula>$H{start_row}="進行中"</formula></cfRule>
<cfRule type='expression' dxfId='6' priority='7'><formula>$H{start_row}="遅延"</formula></cfRule>
<cfRule type='expression' dxfId='7' priority='8'><formula>$H{start_row}="完了"</formula></cfRule>
</conditionalFormatting>
"""

    return [gantt_rules, status_rules]


def case_master_sheet(password_hash: str = "") -> str:
    """Case_Master シートを生成する。

    編集可能: 案件ID(A), 案件名(B), メモ(C) の 2〜100 行目、案件選択(H1)
    保護: 施策数(D), 平均進捗(E), ドリルダウン領域(G3:N104)
    """
    cells: List[Tuple[int, int, object]] = []
    headers = ["案件ID", "案件名", "メモ", "施策数", "平均進捗"]
    for col, header in enumerate(headers, start=1):
        cells.append((1, col, header))

    for idx, (case_id, name) in enumerate(CASES, start=2):
        cells.extend(
            [
                (idx, 1, case_id),
                (idx, 2, name),
                (idx, 4, Formula(f"COUNTIF(Measure_Master!$B:$B,A{idx})")),
                (idx, 5, Formula(f"IFERROR(AVERAGEIF(Measure_Master!$B:$B,A{idx},Measure_Master!$G:$G),0)")),
            ]
        )

    drill_down_headers = [
        "施策ID",
        "親案件ID",
        "施策名",
        "開始日",
        "WBS リンク",
        "WBS シート名",
        "実進捗",
        "備考",
    ]
    for col, header in enumerate(drill_down_headers, start=7):
        cells.append((2, col, header))
    cells.append((1, 7, "案件ドリルダウン"))
    cells.append((1, 8, "CASE-001"))

    cells.append(
        (
            3,
            7,
            Formula(
                "IF($H$1=\"\",\"\",IFERROR(FILTER(MeasureList,INDEX(MeasureList,,2)=$H$1),\"該当なし\"))"
            ),
        )
    )

    data_validations = (
        "<dataValidations count='1'>"
        "<dataValidation type='list' allowBlank='1' showDropDown='1' showErrorMessage='1' showInputMessage='1' errorStyle='stop' errorTitle='入力エラー' error='リストから選択してください' promptTitle='案件IDの選択' prompt='プルダウンから案件IDを選択してください' sqref='H1'>"
        "<formula1>CaseIds</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    # ロック解除セル: A,B,C 列の 2〜100 行目、H1 (案件選択)
    unlocked: set[Tuple[int, int]] = set()
    for row in range(2, 101):
        unlocked.add((row, 1))  # A 列
        unlocked.add((row, 2))  # B 列
        unlocked.add((row, 3))  # C 列
    unlocked.add((1, 8))  # H1

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=False)
    return worksheet_xml(cells, data_validations=data_validations, sheet_protection=protection, unlocked_cells=unlocked)


def measure_master_sheet(password_hash: str = "") -> str:
    """Measure_Master シートを生成する。

    編集可能: 施策ID(A), 親案件ID(B), 施策名(C), 開始日(D), WBSシート名(F), 備考(H) の 2〜104 行目
    保護: WBSリンク(E), 実進捗(G), ヘッダー行
    """
    cells: List[Tuple[int, int, object]] = []
    headers = ["施策ID", "親案件ID", "施策名", "開始日", "WBS リンク", "WBS シート名", "実進捗", "備考"]
    for col, header in enumerate(headers, start=1):
        cells.append((1, col, header))

    for idx, (mid, cid, name, start, sheet_name) in enumerate(MEASURES, start=2):
        cells.extend(
            [
                (idx, 1, mid),
                (idx, 2, cid),
                (idx, 3, name),
                (idx, 4, start),
                (idx, 6, sheet_name),
                (idx, 5, Formula(f"HYPERLINK(\"#'\" & F{idx} & \"'!A1\", \"WBSを開く\")")),
                (idx, 7, Formula(f"IF(F{idx}=\"\",\"\",IFERROR(INDIRECT(\"'\" & F{idx} & \"'!J2\"),\"未リンク\"))")),
            ]
        )

    data_validations = (
        "<dataValidations count='1'>"
        "<dataValidation type='list' allowBlank='0' showDropDown='1' showErrorMessage='1' showInputMessage='1' errorStyle='stop' errorTitle='入力エラー' error='リスト外の値は入力できません' promptTitle='案件IDの選択' prompt='プルダウンから案件IDを選択してください' sqref='B2:B104'>"
        "<formula1>CaseIds</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    # ロック解除セル: A,B,C,D,F,H 列の 2〜104 行目
    unlocked: set[Tuple[int, int]] = set()
    editable_cols = [1, 2, 3, 4, 6, 8]  # A,B,C,D,F,H
    for row in range(2, 105):
        for col in editable_cols:
            unlocked.add((row, col))

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=False)
    return worksheet_xml(cells, data_validations=data_validations, sheet_protection=protection, unlocked_cells=unlocked)


def kanban_sheet(password_hash: str = "") -> str:
    """Kanban_View シートを生成する。

    編集可能: B2 (WBS シート名選択) のみ
    保護: カード生成式 (B5:G104)、ヘッダー (1〜4 行)
    """
    cells: List[Tuple[int, int, object]] = [
        (1, 1, "施策を選択"),
        (1, 2, "WBS シート名"),
        (2, 2, "PRJ_001"),
        (4, 2, "To Do"),
        (4, 4, "Doing"),
        (4, 6, "Done"),
    ]

    formula_template = (
        "IF($B$2=\"\",\"\",IFERROR(LET(_s,$B$2,_tasks,INDIRECT(\"'\"&_s&\"'!B5:B104\"),"
        "_owners,INDIRECT(\"'\"&_s&\"'!C5:C104\"),_due,INDIRECT(\"'\"&_s&\"'!F5:F104\"),"
        "_status,INDIRECT(\"'\"&_s&\"'!H5:H104\"),_filtered,FILTER(HSTACK(_tasks,_owners,_due),_status=\"{status}\"),"
        "MAP(INDEX(_filtered,,1),INDEX(_filtered,,2),INDEX(_filtered,,3),LAMBDA(t,o,d,t&CHAR(10)&o&CHAR(10)&TEXT(d,\"yyyy-mm-dd\"))))),\"選択したWBSシートが見つかりません\"))"
    )

    cells.append((5, 2, Formula(formula_template.format(status="未着手"))))
    cells.append((5, 4, Formula(formula_template.format(status="進行中"))))
    cells.append((5, 6, Formula(formula_template.format(status="完了"))))

    data_validations = (
        "<dataValidations count='1'>"
        "<dataValidation type='list' allowBlank='1' showDropDown='1' showErrorMessage='1' showInputMessage='1' errorStyle='stop' errorTitle='入力エラー' error='リスト外の値は入力できません' promptTitle='WBS シート名の選択' prompt='プルダウンから施策の WBS シート名を選択してください' sqref='B2'>"
        "<formula1>Measure_Master!$F$2:$F$20</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    # ロック解除セル: B2 のみ
    unlocked: set[Tuple[int, int]] = {(2, 2)}

    protection = SheetProtection(password_hash=password_hash, allow_insert_rows=False)
    return worksheet_xml(cells, data_validations=data_validations, sheet_protection=protection, unlocked_cells=unlocked)


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
) -> List[str]:
    """指定した枚数の PRJ シートを生成してブックを書き出し、レポート用テキストを返す。"""

    # パスワードハッシュを計算
    password = get_sheet_password()
    pwd_hash = excel_password_hash(password)

    # Config / Template
    sheet_names = ["Config", "Template"]
    sheets_xml: List[str] = [
        config_sheet(password_hash=pwd_hash),
        template_sheet(sample=False, password_hash=pwd_hash),
    ]

    # PRJ_xxx をまとめて生成
    for idx in range(1, project_count + 1):
        sheet_names.append(f"PRJ_{idx:03d}")
        is_sample = sample_all_projects or (sample_first_project and idx == 1)
        sheets_xml.append(template_sheet(sample=is_sample, password_hash=pwd_hash))

    # 末尾のマスターシート群
    sheet_names.extend(["Case_Master", "Measure_Master", "Kanban_View"])
    sheets_xml.extend([
        case_master_sheet(password_hash=pwd_hash),
        measure_master_sheet(password_hash=pwd_hash),
        kanban_sheet(password_hash=pwd_hash),
    ])

    defined_names = {
        "CaseIds": "Case_Master!$A$2:$A$100",
        "MeasureList": "Measure_Master!$A$2:$H$104",
        "CaseDrilldownArea": "Case_Master!$G$3:$N$104",
    }

    vba_modules = load_vba_modules()
    vba_binary = vba_project_binary(vba_modules)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets_xml)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names, defined_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets_xml)))
        zf.writestr("xl/styles.xml", styles_xml())
        zf.writestr("xl/vbaProject.bin", vba_binary)

        for idx, xml in enumerate(sheets_xml, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", xml)

    print(f"ブックを生成しました: {output_path}")

    return generate_report_lines(project_count, sample_first_project, sample_all_projects, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Modern Excel PMS 雛形を生成する")
    parser.add_argument("--projects", type=int, default=1, help="生成する PRJ_xxx シート数")
    parser.add_argument(
        "--sample-first",
        action="store_true",
        help="最初の PRJ シートにサンプルタスクを埋め込む",
    )
    parser.add_argument(
        "--sample-all",
        action="store_true",
        help="全ての PRJ シートにサンプルタスクを埋め込む",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_PATH,
        help="出力先パス (.xlsm)",
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
    report_lines = build_workbook(args.projects, args.sample_first, args.sample_all, args.output)

    if args.report_output:
        write_report_text(report_lines, args.report_output)
        print(f"レポートを出力しました: {args.report_output}")

    if args.pdf_output:
        export_pdf_report(report_lines, args.pdf_output)
        print(f"PDF レポートを出力しました: {args.pdf_output}")


if __name__ == "__main__":
    main()
