#!/usr/bin/env python3
"""
Excel WBS ファイルの仕様検証テストスクリプト

README.mdの仕様に基づいてxlsxファイルを検証する。
"""
import os
import sys
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
import re

# OpenXML名前空間
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}


@dataclass
class ValidationResult:
    """検証結果"""
    passed: bool
    message: str
    details: Optional[str] = None


@dataclass
class TestReport:
    """テストレポート"""
    test_name: str
    results: List[ValidationResult]

    @property
    def passed(self) -> bool:
        return all(r.passed for r in self.results)

    @property
    def passed_count(self) -> int:
        return sum(1 for r in self.results if r.passed)

    @property
    def failed_count(self) -> int:
        return sum(1 for r in self.results if not r.passed)


class XlsxValidator:
    """Excelファイルの検証クラス"""

    def __init__(self, xlsx_path: str):
        self.xlsx_path = xlsx_path
        self.temp_dir = None
        self.extracted_path = None

    def __enter__(self):
        self.temp_dir = tempfile.mkdtemp(prefix='xlsx_validate_')
        with zipfile.ZipFile(self.xlsx_path, 'r') as zf:
            zf.extractall(self.temp_dir)
        self.extracted_path = Path(self.temp_dir)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        import shutil
        if self.temp_dir:
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def read_xml(self, relative_path: str) -> Optional[ET.Element]:
        """XMLファイルを読み込む"""
        full_path = self.extracted_path / relative_path
        if not full_path.exists():
            return None
        try:
            tree = ET.parse(full_path)
            return tree.getroot()
        except ET.ParseError as e:
            print(f"XML Parse Error in {relative_path}: {e}")
            return None

    def get_sheet_names(self) -> List[str]:
        """シート名一覧を取得"""
        workbook = self.read_xml('xl/workbook.xml')
        if workbook is None:
            return []

        sheets = workbook.findall('.//main:sheet', NS)
        return [s.get('name', '') for s in sheets]

    def get_sheet_xml(self, sheet_index: int) -> Optional[ET.Element]:
        """シートのXMLを取得"""
        return self.read_xml(f'xl/worksheets/sheet{sheet_index + 1}.xml')

    def get_styles(self) -> Optional[ET.Element]:
        """styles.xmlを取得"""
        return self.read_xml('xl/styles.xml')


def validate_file_structure(validator: XlsxValidator) -> TestReport:
    """ファイル構造の検証"""
    results = []

    # 必須ファイルの存在確認
    required_files = [
        '[Content_Types].xml',
        '_rels/.rels',
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/styles.xml',
    ]

    for f in required_files:
        path = validator.extracted_path / f
        if path.exists():
            results.append(ValidationResult(True, f"必須ファイル存在: {f}"))
        else:
            results.append(ValidationResult(False, f"必須ファイル不足: {f}"))

    return TestReport("ファイル構造検証", results)


def validate_sheet_names(validator: XlsxValidator) -> TestReport:
    """シート名の検証"""
    results = []

    sheet_names = validator.get_sheet_names()

    # 必須シート
    required_sheets = ['Config', 'Template']
    for s in required_sheets:
        if s in sheet_names:
            results.append(ValidationResult(True, f"必須シート存在: {s}"))
        else:
            results.append(ValidationResult(False, f"必須シート不足: {s}"))

    # シート数の確認
    if len(sheet_names) >= 2:
        results.append(ValidationResult(True, f"シート数: {len(sheet_names)}"))
    else:
        results.append(ValidationResult(False, f"シート数不足: {len(sheet_names)}"))

    return TestReport("シート名検証", results)


def validate_styles(validator: XlsxValidator) -> TestReport:
    """スタイルの検証"""
    results = []

    styles = validator.get_styles()
    if styles is None:
        results.append(ValidationResult(False, "styles.xml読み込み失敗"))
        return TestReport("スタイル検証", results)

    # フォントの検証
    fonts = styles.findall('.//main:font', NS)
    font_names = []
    for font in fonts:
        name_elem = font.find('main:name', NS)
        if name_elem is not None:
            font_names.append(name_elem.get('val', ''))

    if 'Meiryo UI' in font_names:
        results.append(ValidationResult(True, "Meiryo UIフォント使用"))
    else:
        results.append(ValidationResult(False, f"Meiryo UIフォント未使用 (fonts: {font_names})"))

    # 塗りつぶしの検証 (ヘッダー色 #2C3E50)
    fills = styles.findall('.//main:fill', NS)
    has_header_fill = False
    for fill in fills:
        fg_color = fill.find('.//main:fgColor', NS)
        if fg_color is not None:
            color = fg_color.get('rgb', '').upper()
            if '2C3E50' in color:
                has_header_fill = True
                break

    if has_header_fill:
        results.append(ValidationResult(True, "ヘッダー背景色 #2C3E50 使用"))
    else:
        results.append(ValidationResult(False, "ヘッダー背景色 #2C3E50 未使用"))

    # 入力セル用の薄い青色の確認
    has_input_fill = False
    for fill in fills:
        fg_color = fill.find('.//main:fgColor', NS)
        if fg_color is not None:
            color = fg_color.get('rgb', '').upper()
            if 'EAF2F8' in color or 'D5E8F7' in color:
                has_input_fill = True
                break

    if has_input_fill:
        results.append(ValidationResult(True, "入力セル背景色（薄青）使用"))
    else:
        results.append(ValidationResult(False, "入力セル背景色（薄青）未使用"))

    # cellXfsの数を確認
    cell_xfs = styles.find('.//main:cellXfs', NS)
    if cell_xfs is not None:
        xf_count = int(cell_xfs.get('count', '0'))
        if xf_count >= 5:
            results.append(ValidationResult(True, f"セルスタイル数: {xf_count}"))
        else:
            results.append(ValidationResult(False, f"セルスタイル数不足: {xf_count}"))

    # 条件付き書式（dxfs）の検証
    dxfs = styles.find('.//main:dxfs', NS)
    if dxfs is not None:
        dxf_count = int(dxfs.get('count', '0'))
        if dxf_count >= 4:
            results.append(ValidationResult(True, f"条件付き書式スタイル数: {dxf_count}"))
        else:
            results.append(ValidationResult(False, f"条件付き書式スタイル数不足: {dxf_count}"))

    return TestReport("スタイル検証", results)


def validate_template_sheet(validator: XlsxValidator) -> TestReport:
    """Templateシート（WBS）の検証"""
    results = []

    # Templateシートのインデックスを取得
    sheet_names = validator.get_sheet_names()
    if 'Template' not in sheet_names:
        results.append(ValidationResult(False, "Templateシートが存在しない"))
        return TestReport("Templateシート検証", results)

    template_idx = sheet_names.index('Template')
    sheet = validator.get_sheet_xml(template_idx)

    if sheet is None:
        results.append(ValidationResult(False, "Templateシート読み込み失敗"))
        return TestReport("Templateシート検証", results)

    # 列幅の設定確認
    cols = sheet.find('.//main:cols', NS)
    if cols is not None:
        col_count = len(cols.findall('main:col', NS))
        if col_count >= 5:
            results.append(ValidationResult(True, f"列幅設定数: {col_count}"))
        else:
            results.append(ValidationResult(False, f"列幅設定不足: {col_count}"))
    else:
        results.append(ValidationResult(False, "列幅設定なし"))

    # フリーズペインの確認
    pane = sheet.find('.//main:pane', NS)
    if pane is not None:
        x_split = pane.get('xSplit', '0')
        y_split = pane.get('ySplit', '0')
        state = pane.get('state', '')
        if state == 'frozen' and int(y_split) > 0:
            results.append(ValidationResult(True, f"フリーズペイン設定: 行{y_split}, 列{x_split}"))
        else:
            results.append(ValidationResult(False, f"フリーズペイン不適切: state={state}"))
    else:
        results.append(ValidationResult(False, "フリーズペイン設定なし"))

    # ヘッダー行（4行目）の確認
    sheet_data = sheet.find('.//main:sheetData', NS)
    if sheet_data is not None:
        rows = sheet_data.findall('main:row', NS)
        row4 = None
        for row in rows:
            if row.get('r') == '4':
                row4 = row
                break

        if row4 is not None:
            cells = row4.findall('main:c', NS)
            header_cells = []
            for cell in cells:
                inline_str = cell.find('.//main:t', NS)
                if inline_str is not None:
                    header_cells.append(inline_str.text)

            expected_headers = ['Lv', 'タスク名', '担当', '開始日', '工数', '終了日', '進捗率', 'ステータス']
            found_headers = []
            for h in expected_headers:
                for hc in header_cells:
                    if hc and h in hc:
                        found_headers.append(h)
                        break

            if len(found_headers) >= 6:
                results.append(ValidationResult(True, f"ヘッダー列: {len(found_headers)}/8"))
            else:
                results.append(ValidationResult(False, f"ヘッダー列不足: {found_headers}"))
        else:
            results.append(ValidationResult(False, "ヘッダー行(4行目)が見つからない"))

    # 条件付き書式の確認
    cf = sheet.findall('.//main:conditionalFormatting', NS)
    if len(cf) >= 2:
        results.append(ValidationResult(True, f"条件付き書式セクション数: {len(cf)}"))
    else:
        results.append(ValidationResult(False, f"条件付き書式不足: {len(cf)}"))

    # ガントチャート用の条件付き書式ルールの確認
    gantt_rules = 0
    status_rules = 0
    for cf_section in cf:
        sqref = cf_section.get('sqref', '')
        rules = cf_section.findall('main:cfRule', NS)

        # ガントチャート範囲 (K列以降)
        if sqref and ('K5' in sqref or 'K5:' in sqref):
            gantt_rules = len(rules)
        # ステータス列 (H列)
        if sqref and 'H5' in sqref:
            status_rules = len(rules)

    if gantt_rules >= 3:
        results.append(ValidationResult(True, f"ガントチャート条件付き書式: {gantt_rules}ルール"))
    else:
        results.append(ValidationResult(False, f"ガントチャート条件付き書式不足: {gantt_rules}ルール"))

    if status_rules >= 3:
        results.append(ValidationResult(True, f"ステータス条件付き書式: {status_rules}ルール"))
    else:
        results.append(ValidationResult(False, f"ステータス条件付き書式不足: {status_rules}ルール"))

    # シート保護の確認
    protection = sheet.find('.//main:sheetProtection', NS)
    if protection is not None:
        results.append(ValidationResult(True, "シート保護設定あり"))
    else:
        results.append(ValidationResult(False, "シート保護設定なし"))

    # データバリデーションの確認
    dv = sheet.find('.//main:dataValidations', NS)
    if dv is not None:
        dv_count = int(dv.get('count', '0'))
        if dv_count >= 1:
            results.append(ValidationResult(True, f"データバリデーション: {dv_count}"))
        else:
            results.append(ValidationResult(False, "データバリデーションなし"))
    else:
        results.append(ValidationResult(False, "データバリデーション設定なし"))

    return TestReport("Templateシート検証", results)


def validate_config_sheet(validator: XlsxValidator) -> TestReport:
    """Configシートの検証"""
    results = []

    sheet_names = validator.get_sheet_names()
    if 'Config' not in sheet_names:
        results.append(ValidationResult(False, "Configシートが存在しない"))
        return TestReport("Configシート検証", results)

    config_idx = sheet_names.index('Config')
    sheet = validator.get_sheet_xml(config_idx)

    if sheet is None:
        results.append(ValidationResult(False, "Configシート読み込み失敗"))
        return TestReport("Configシート検証", results)

    # 必須項目の確認
    sheet_data = sheet.find('.//main:sheetData', NS)
    if sheet_data is None:
        results.append(ValidationResult(False, "sheetDataなし"))
        return TestReport("Configシート検証", results)

    # 全セルのテキストを収集
    all_text = []
    for row in sheet_data.findall('main:row', NS):
        for cell in row.findall('main:c', NS):
            t = cell.find('.//main:t', NS)
            if t is not None and t.text:
                all_text.append(t.text)

    # 必須ラベルの確認
    required_labels = ['祝日リスト', '担当者リスト', 'ステータスリスト']
    for label in required_labels:
        found = any(label in text for text in all_text)
        if found:
            results.append(ValidationResult(True, f"必須ラベル存在: {label}"))
        else:
            results.append(ValidationResult(False, f"必須ラベル不足: {label}"))

    return TestReport("Configシート検証", results)


def validate_formulas(validator: XlsxValidator) -> TestReport:
    """数式の検証"""
    results = []

    sheet_names = validator.get_sheet_names()
    if 'Template' not in sheet_names:
        results.append(ValidationResult(False, "Templateシートが存在しない"))
        return TestReport("数式検証", results)

    template_idx = sheet_names.index('Template')
    sheet = validator.get_sheet_xml(template_idx)

    if sheet is None:
        results.append(ValidationResult(False, "Templateシート読み込み失敗"))
        return TestReport("数式検証", results)

    # 数式を収集
    formulas = []
    sheet_data = sheet.find('.//main:sheetData', NS)
    if sheet_data is not None:
        for row in sheet_data.findall('main:row', NS):
            for cell in row.findall('main:c', NS):
                f = cell.find('main:f', NS)
                if f is not None and f.text:
                    formulas.append((cell.get('r', ''), f.text))

    # 終了日計算（WORKDAY関数）
    workday_found = any('WORKDAY' in f[1] for f in formulas)
    if workday_found:
        results.append(ValidationResult(True, "WORKDAY関数使用"))
    else:
        results.append(ValidationResult(False, "WORKDAY関数未使用"))

    # ステータス自動計算（IFS関数）
    ifs_found = any('IFS' in f[1] for f in formulas)
    if ifs_found:
        results.append(ValidationResult(True, "IFS関数使用（ステータス自動計算）"))
    else:
        results.append(ValidationResult(False, "IFS関数未使用"))

    # 全体進捗計算（LET関数またはSUMPRODUCT）
    progress_found = any('LET' in f[1] or 'SUMPRODUCT' in f[1] for f in formulas)
    if progress_found:
        results.append(ValidationResult(True, "全体進捗計算数式あり"))
    else:
        results.append(ValidationResult(False, "全体進捗計算数式なし"))

    # TODAY関数（ガント基準日）
    today_found = any('TODAY' in f[1] for f in formulas)
    if today_found:
        results.append(ValidationResult(True, "TODAY関数使用"))
    else:
        results.append(ValidationResult(False, "TODAY関数未使用"))

    # 数式の総数
    if len(formulas) >= 10:
        results.append(ValidationResult(True, f"数式総数: {len(formulas)}"))
    else:
        results.append(ValidationResult(False, f"数式数不足: {len(formulas)}"))

    return TestReport("数式検証", results)


def validate_cell_styles_applied(validator: XlsxValidator) -> TestReport:
    """セルスタイルの適用状況を検証"""
    results = []

    sheet_names = validator.get_sheet_names()
    if 'Template' not in sheet_names:
        results.append(ValidationResult(False, "Templateシートが存在しない"))
        return TestReport("セルスタイル適用検証", results)

    template_idx = sheet_names.index('Template')
    sheet = validator.get_sheet_xml(template_idx)

    if sheet is None:
        results.append(ValidationResult(False, "Templateシート読み込み失敗"))
        return TestReport("セルスタイル適用検証", results)

    # スタイル適用状況を収集
    style_usage = {}
    sheet_data = sheet.find('.//main:sheetData', NS)
    if sheet_data is not None:
        for row in sheet_data.findall('main:row', NS):
            for cell in row.findall('main:c', NS):
                s = cell.get('s', '0')
                style_usage[s] = style_usage.get(s, 0) + 1

    # スタイル0以外が使われているか
    non_default_styles = {k: v for k, v in style_usage.items() if k != '0'}
    if non_default_styles:
        results.append(ValidationResult(True, f"カスタムスタイル使用: {len(non_default_styles)}種類"))
    else:
        results.append(ValidationResult(False, "カスタムスタイル未使用（すべてデフォルト）"))

    # ヘッダー行（4行目）のスタイル確認
    row4_styles = set()
    if sheet_data is not None:
        for row in sheet_data.findall('main:row', NS):
            if row.get('r') == '4':
                for cell in row.findall('main:c', NS):
                    s = cell.get('s', '0')
                    row4_styles.add(s)

    if len(row4_styles) > 0 and '0' not in row4_styles:
        results.append(ValidationResult(True, f"ヘッダー行スタイル適用: {row4_styles}"))
    elif '2' in row4_styles:
        results.append(ValidationResult(True, f"ヘッダー行にヘッダースタイル(2)適用"))
    else:
        results.append(ValidationResult(False, f"ヘッダー行スタイル問題: {row4_styles}"))

    # 入力セル（5行目以降）のスタイル確認
    input_styles = set()
    if sheet_data is not None:
        for row in sheet_data.findall('main:row', NS):
            row_num = int(row.get('r', '0'))
            if row_num >= 5:
                for cell in row.findall('main:c', NS):
                    s = cell.get('s', '0')
                    input_styles.add(s)

    if input_styles and ('3' in input_styles or '9' in input_styles):
        results.append(ValidationResult(True, f"入力セルスタイル適用: {input_styles}"))
    else:
        results.append(ValidationResult(False, f"入力セルスタイル問題: {input_styles}"))

    return TestReport("セルスタイル適用検証", results)


def validate_date_values(validator: XlsxValidator) -> TestReport:
    """日付値の検証（シリアル値であることを確認）

    PRJ_001 シート（サンプルデータあり）を対象に検証する。
    """
    results = []

    sheet_names = validator.get_sheet_names()

    # PRJ_001 シートを優先して検証（サンプルデータがある）
    target_sheet = None
    target_name = None
    for name in ['PRJ_001', 'Template']:
        if name in sheet_names:
            target_name = name
            target_sheet = validator.get_sheet_xml(sheet_names.index(name))
            break

    if target_sheet is None:
        results.append(ValidationResult(False, "PRJ_001またはTemplateシートが存在しない"))
        return TestReport("日付値検証", results)

    sheet = target_sheet

    if sheet is None:
        results.append(ValidationResult(False, f"{target_name}シート読み込み失敗"))
        return TestReport("日付値検証", results)

    results.append(ValidationResult(True, f"検証対象シート: {target_name}"))

    # D列（開始日）の値を確認
    date_cells = []
    sheet_data = sheet.find('.//main:sheetData', NS)
    if sheet_data is not None:
        for row in sheet_data.findall('main:row', NS):
            row_num = int(row.get('r', '0'))
            if row_num >= 5:  # データ行
                for cell in row.findall('main:c', NS):
                    cell_ref = cell.get('r', '')
                    if cell_ref.startswith('D'):
                        v = cell.find('main:v', NS)
                        inline_str = cell.find('.//main:t', NS)

                        if v is not None:
                            try:
                                val = float(v.text)
                                if 40000 < val < 50000:  # Excel日付範囲
                                    date_cells.append((cell_ref, 'numeric', val))
                                else:
                                    date_cells.append((cell_ref, 'numeric_other', val))
                            except (ValueError, TypeError):
                                date_cells.append((cell_ref, 'text', v.text))
                        elif inline_str is not None and inline_str.text:
                            # 空のinline_stringは無視（スタイルのみのセル）
                            date_cells.append((cell_ref, 'inline_string', inline_str.text))

    numeric_dates = [d for d in date_cells if d[1] == 'numeric']
    string_dates = [d for d in date_cells if d[1] in ('text', 'inline_string')]

    if numeric_dates:
        results.append(ValidationResult(True, f"数値形式の日付: {len(numeric_dates)}セル"))
    else:
        results.append(ValidationResult(False, "数値形式の日付なし"))

    if string_dates:
        results.append(ValidationResult(False, f"文字列形式の日付あり: {len(string_dates)}セル ({string_dates[:3]})"))
    else:
        results.append(ValidationResult(True, "文字列形式の日付なし"))

    return TestReport("日付値検証", results)


def run_all_validations(xlsx_path: str) -> List[TestReport]:
    """全検証を実行"""
    reports = []

    with XlsxValidator(xlsx_path) as validator:
        reports.append(validate_file_structure(validator))
        reports.append(validate_sheet_names(validator))
        reports.append(validate_styles(validator))
        reports.append(validate_config_sheet(validator))
        reports.append(validate_template_sheet(validator))
        reports.append(validate_formulas(validator))
        reports.append(validate_cell_styles_applied(validator))
        reports.append(validate_date_values(validator))

    return reports


def print_report(reports: List[TestReport]):
    """レポートを出力"""
    total_passed = 0
    total_failed = 0

    print("\n" + "=" * 70)
    print("Excel WBS ファイル検証レポート")
    print("=" * 70)

    for report in reports:
        status = "✅ PASS" if report.passed else "❌ FAIL"
        print(f"\n## {report.test_name} [{status}]")
        print("-" * 50)

        for result in report.results:
            icon = "✓" if result.passed else "✗"
            print(f"  {icon} {result.message}")
            if result.details:
                print(f"      {result.details}")

        total_passed += report.passed_count
        total_failed += report.failed_count

    print("\n" + "=" * 70)
    print(f"総合結果: {total_passed} passed, {total_failed} failed")

    if total_failed == 0:
        print("🎉 すべてのテストに合格しました！")
    else:
        print(f"⚠️  {total_failed}件の問題が見つかりました")

    print("=" * 70)

    return total_failed == 0


def main():
    if len(sys.argv) < 2:
        # デフォルトのパス
        xlsx_path = '/home/ec2-user/workspace/SUPER-WBS-IN-EXCEL/output/ModernExcelPMS.xlsx'
    else:
        xlsx_path = sys.argv[1]

    if not os.path.exists(xlsx_path):
        print(f"Error: File not found: {xlsx_path}")
        sys.exit(1)

    print(f"検証対象: {xlsx_path}")

    reports = run_all_validations(xlsx_path)
    success = print_report(reports)

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
