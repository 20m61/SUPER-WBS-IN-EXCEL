"""
Modern Excel PMS の雛形ブックを自動生成するスクリプト。
外部ライブラリに依存せず、OpenXML を直接書き出して `ModernExcelPMS.xlsx` を生成する。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Sequence, Tuple
from xml.sax.saxutils import escape
import zipfile

OUTPUT_PATH = Path(__file__).resolve().parent.parent / "ModernExcelPMS.xlsx"


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
        "ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>"
        f"{overrides}"
        "<Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/>"
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
        "<dxfs count='0'/><tableStyles count='0' defaultTableStyle='TableStyleMedium9' defaultPivotStyle='PivotStyleLight16'/>"
        "</styleSheet>"
    )


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
    return worksheet_xml(template_cells(sample), data_validations=template_data_validations())


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
        "<dataValidation type='list' allowBlank='1' showDropDown='1' sqref='H1'>"
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
                (idx, 7, Formula(f"INDIRECT(\"'\" & F{idx} & \"'!J2\")")),
            ]
        )

    data_validations = (
        "<dataValidations count='1'>"
        "<dataValidation type='list' allowBlank='1' showDropDown='1' sqref='B2:B104'>"
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
        "IF($B$2=\"\",\"\",LET(_s,$B$2,_tasks,INDIRECT(\"'\"&_s&\"'!B5:B104\"),"
        "_owners,INDIRECT(\"'\"&_s&\"'!C5:C104\"),_due,INDIRECT(\"'\"&_s&\"'!F5:F104\"),"
        "_status,INDIRECT(\"'\"&_s&\"'!H5:H104\"),_filtered,FILTER(HSTACK(_tasks,_owners,_due),_status=\"{status}\"),"
        "MAP(INDEX(_filtered,,1),INDEX(_filtered,,2),INDEX(_filtered,,3),LAMBDA(t,o,d,t&CHAR(10)&o&CHAR(10)&TEXT(d,\"yyyy-mm-dd\")))))"
    )

    cells.append((5, 2, Formula(formula_template.format(status="未着手"))))
    cells.append((5, 4, Formula(formula_template.format(status="進行中"))))
    cells.append((5, 6, Formula(formula_template.format(status="完了"))))

    data_validations = (
        "<dataValidations count='1'>"
        "<dataValidation type='list' allowBlank='1' showDropDown='1' sqref='B2'>"
        "<formula1>Measure_Master!$F$2:$F$20</formula1>"
        "</dataValidation>"
        "</dataValidations>"
    )

    return worksheet_xml(cells, data_validations=data_validations)


# --------------------------- メイン ---------------------------

def build_workbook() -> None:
    sheet_builders = [
        config_sheet,
        lambda: template_sheet(sample=False),
        lambda: template_sheet(sample=True),
        case_master_sheet,
        measure_master_sheet,
        kanban_sheet,
    ]
    sheet_names = [
        "Config",
        "Template",
        "PRJ_001",
        "Case_Master",
        "Measure_Master",
        "Kanban_View",
    ]

    sheets_xml = [builder() for builder in sheet_builders]

    with zipfile.ZipFile(OUTPUT_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets_xml)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets_xml)))
        zf.writestr("xl/styles.xml", styles_xml())

        for idx, xml in enumerate(sheets_xml, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", xml)

    print(f"ブックを生成しました: {OUTPUT_PATH}")


if __name__ == "__main__":
    build_workbook()
