"""
Modern Excel PMS の雛形ブックを自動生成するスクリプト。
外部ライブラリに依存せず、OpenXML を直接書き出して `ModernExcelPMS.xlsm` を生成する。
"""

from __future__ import annotations

from dataclasses import dataclass
import argparse
from pathlib import Path
from typing import List, Mapping, Sequence, Tuple
from xml.sax.saxutils import escape
import zipfile

OUTPUT_PATH = Path(__file__).resolve().parent.parent / "ModernExcelPMS.xlsm"
VBA_SOURCE_DIR = Path(__file__).resolve().parent.parent / "docs" / "vba"


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


def cell_xml(row: int, col: int, value) -> str:
    ref = cell_ref(row, col)
    if isinstance(value, Formula):
        return f"<c r=\"{ref}\"><f>{escape(value.expr)}</f></c>"
    if isinstance(value, str):
        return f"<c r=\"{ref}\" t=\"inlineStr\"><is><t>{escape(value)}</t></is></c>"
    if value is None:
        return ""
    return f"<c r=\"{ref}\"><v>{value}</v></c>"


def worksheet_xml(
    cells: Sequence[Tuple[int, int, object]],
    data_validations: str | None = None,
    conditional_formattings: Sequence[str] | None = None,
) -> str:
    rows = {}
    for row, col, value in cells:
        rows.setdefault(row, {})[col] = value

    xml_lines: List[str] = [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheetData>",
    ]

    for row_idx in sorted(rows):
        xml_lines.append(f"<row r=\"{row_idx}\">")
        for col_idx in sorted(rows[row_idx]):
            xml_lines.append(cell_xml(row_idx, col_idx, rows[row_idx][col_idx]))
        xml_lines.append("</row>")

    xml_lines.append("</sheetData>")

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


def workbook_xml(sheet_names: Sequence[str]) -> str:
    sheets_xml = "".join(
        f"<sheet name='{escape(name)}' sheetId='{idx}' r:id='rId{idx}'/>"
        for idx, name in enumerate(sheet_names, start=1)
    )
    return (
        "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>"
        f"<sheets>{sheets_xml}</sheets>"
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
    return (
        "<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
        "<fonts count='1'><font><sz val='11'/><color theme='1'/><name val='Calibri'/><family val='2'/><scheme val='minor'/></font></fonts>"
        "<fills count='2'><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill></fills>"
        "<borders count='1'><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
        "<cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs>"
        "<cellXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/></cellXfs>"
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

def config_sheet() -> str:
    cells = [
        (1, 1, "祝日リスト"),
        (2, 1, "日付"),
        (1, 4, "担当者マスタ"),
        (2, 4, "氏名"),
        (1, 6, "ステータス候補"),
        (2, 6, "値"),
    ]
    holidays = ["2024-01-01", "2024-02-12", "2024-04-29", "2024-05-03", "2024-05-04", "2024-05-05"]
    for idx, day in enumerate(holidays, start=4):
        cells.append((idx, 2, day))
    members = ["PM_佐藤", "TL_田中", "DEV_鈴木", "QA_伊藤"]
    for idx, member in enumerate(members, start=4):
        cells.append((idx, 4, member))
    statuses = ["未着手", "進行中", "遅延", "完了"]
    for idx, status in enumerate(statuses, start=4):
        cells.append((idx, 6, status))
    return worksheet_xml(cells)


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
        cells.extend(
            [
                (5, 1, 1),
                (5, 2, "キックオフ準備"),
                (5, 3, "PM_佐藤"),
                (5, 4, "2024-05-07"),
                (5, 5, 2),
                (5, 7, 1),
                (6, 1, 2),
                (6, 2, "要件定義ワークショップ"),
                (6, 3, "TL_田中"),
                (6, 4, "2024-05-09"),
                (6, 5, 3),
                (6, 7, 0.5),
                (7, 1, 2),
                (7, 2, "WBS 詳細化"),
                (7, 3, "DEV_鈴木"),
                (7, 4, "2024-05-13"),
                (7, 5, 5),
                (7, 7, 0),
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


def template_sheet(sample: bool = False) -> str:
    return worksheet_xml(
        template_cells(sample),
        data_validations=template_data_validations(),
        conditional_formattings=template_conditional_formattings(),
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


def case_master_sheet() -> str:
    cells: List[Tuple[int, int, object]] = []
    headers = ["案件ID", "案件名", "メモ", "施策数", "平均進捗"]
    for col, header in enumerate(headers, start=1):
        cells.append((1, col, header))

    cases = [("CASE-001", "Web 刷新案件"), ("CASE-002", "新規 SFA 導入")]
    for idx, (case_id, name) in enumerate(cases, start=2):
        cells.extend(
            [
                (idx, 1, case_id),
                (idx, 2, name),
                (idx, 4, Formula(f"COUNTIF(Measure_Master!$B:$B,A{idx})")),
                (idx, 5, Formula(f"IFERROR(AVERAGEIF(Measure_Master!$B:$B,A{idx},Measure_Master!$G:$G),0)")),
            ]
        )

    drill_down_headers = ["案件選択", "施策ID", "親案件ID", "施策名", "開始日", "WBS リンク", "WBS シート名", "実進捗"]
    for col, header in enumerate(drill_down_headers, start=7):
        cells.append((2, col, header))
    cells.append((1, 7, "案件ドリルダウン"))
    cells.append((1, 8, "CASE-001"))

    cells.append(
        (
            3,
            7,
            Formula(
                "IF($H$1=\"\",\"\",FILTER(Measure_Master!A2:G104,Measure_Master!B2:B104=$H$1,\"該当なし\"))"
            ),
        )
    )

    data_validations = (
        "<dataValidations count='1'>"
        "<dataValidation type='list' allowBlank='1' showDropDown='1' showErrorMessage='1' showInputMessage='1' errorStyle='stop' errorTitle='入力エラー' error='リストから選択してください' promptTitle='案件IDの選択' prompt='プルダウンから案件IDを選択してください' sqref='H1'>"
        "<formula1>Case_Master!$A$2:$A$100</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    return worksheet_xml(cells, data_validations=data_validations)


def measure_master_sheet() -> str:
    cells: List[Tuple[int, int, object]] = []
    headers = ["施策ID", "親案件ID", "施策名", "開始日", "WBS リンク", "WBS シート名", "実進捗", "備考"]
    for col, header in enumerate(headers, start=1):
        cells.append((1, col, header))

    measures = [
        ("ME-001", "CASE-001", "LP 改修", "2024-05-07", "PRJ_001"),
        ("ME-002", "CASE-002", "SFA 導入 PoC", "2024-05-13", "PRJ_002"),
    ]

    for idx, (mid, cid, name, start, sheet_name) in enumerate(measures, start=2):
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
        "<formula1>Case_Master!$A$2:$A$100</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    return worksheet_xml(cells, data_validations=data_validations)


def kanban_sheet() -> str:
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

    return worksheet_xml(cells, data_validations=data_validations)


# --------------------------- メイン ---------------------------

def build_workbook(
    project_count: int,
    sample_first_project: bool,
    sample_all_projects: bool,
    output_path: Path,
) -> None:
    """指定した枚数の PRJ シートを生成してブックを書き出す。"""

    # Config / Template
    sheet_names = ["Config", "Template"]
    sheets_xml: List[str] = [config_sheet(), template_sheet(sample=False)]

    # PRJ_xxx をまとめて生成
    for idx in range(1, project_count + 1):
        sheet_names.append(f"PRJ_{idx:03d}")
        is_sample = sample_all_projects or (sample_first_project and idx == 1)
        sheets_xml.append(template_sheet(sample=is_sample))

    # 末尾のマスターシート群
    sheet_names.extend(["Case_Master", "Measure_Master", "Kanban_View"])
    sheets_xml.extend([case_master_sheet(), measure_master_sheet(), kanban_sheet()])

    vba_modules = load_vba_modules()
    vba_binary = vba_project_binary(vba_modules)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets_xml)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets_xml)))
        zf.writestr("xl/styles.xml", styles_xml())
        zf.writestr("xl/vbaProject.bin", vba_binary)

        for idx, xml in enumerate(sheets_xml, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", xml)

    print(f"ブックを生成しました: {output_path}")


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
    args = parser.parse_args()
    build_workbook(args.projects, args.sample_first, args.sample_all, args.output)


if __name__ == "__main__":
    main()
