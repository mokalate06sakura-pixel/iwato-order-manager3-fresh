import math
import re
from copy import copy
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd
from openpyxl.cell.cell import MergedCell

# 出力する北部市場発注書の書体を統一する
ORDER_FONT_NAME = "游ゴシック"
ORDER_BODY_FONT_SIZE = 11
ORDER_TITLE_FONT_SIZE = 14
ORDER_DATE_FONT_SIZE = 12


def _standardize_order_fonts(ws, is_tokuyou: bool):
    """罫線・寸法・印刷設定を保ち、出力シートの書体と配置を揃える。"""
    date_address = TOKUYOU_DELIVERY_CELL if is_tokuyou else YUHOUSE_DELIVERY_CELL
    quantity_columns = (4, 5, 6) if is_tokuyou else (4,)
    for row in ws:
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            font = copy(cell.font)
            font.name = ORDER_FONT_NAME
            font.scheme = None
            font.sz = ORDER_BODY_FONT_SIZE
            font.bold = False
            text = re.sub(r"\s+", "", str(cell.value or ""))
            if cell.coordinate == date_address:
                font.sz = ORDER_DATE_FONT_SIZE
                font.bold = True
            elif cell.row < DETAIL_START_ROW and (
                SUPPLIER_NAME in text or "御中" in text or "仕入先" in text
            ):
                font.sz = ORDER_DATE_FONT_SIZE
                font.bold = True
            elif cell.row <= 3 and any(
                word in text for word in ("発注", "特養", "ユーハウス", "施設名")
            ):
                font.sz = ORDER_TITLE_FONT_SIZE
                font.bold = True
            elif cell.row == DETAIL_START_ROW - 1:
                font.bold = True
            cell.font = font

            alignment = copy(cell.alignment)
            if cell.row == DETAIL_START_ROW - 1:
                alignment.horizontal = "center"
                alignment.vertical = "center"
            elif DETAIL_START_ROW <= cell.row <= DETAIL_END_ROW:
                alignment.vertical = "center"
                if cell.column in (1, 2, *quantity_columns):
                    alignment.horizontal = "center"
                if cell.column == 2:
                    # 11ptの指定と行高・列幅を保ち、長い食品名をセル内に収める。
                    alignment.wrap_text = False
                    alignment.shrink_to_fit = True
                if cell.column in quantity_columns:
                    cell.number_format = "General"
            cell.alignment = alignment


SUPPLIER_NAME = "北部市場販売"

TOKUYOU_SHEET = "特養 (北部市場)"
YUHOUSE_SHEET = "ユーハウス(北部市場)"

DETAIL_START_ROW = 7
DETAIL_END_ROW = 18
ROWS_PER_PAGE = DETAIL_END_ROW - DETAIL_START_ROW + 1

TOKUYOU_DELIVERY_CELL = "J4"
YUHOUSE_DELIVERY_CELL = "I4"


def _norm(s) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\u3000", " ")).strip()


def _parse_mmdd(value):
    if value is None:
        return None
    s = str(value)
    m = re.search(r"(\d{1,2})/(\d{1,2})", s)
    if not m:
        return None
    mm = int(m.group(1))
    dd = int(m.group(2))
    return pd.Timestamp(year=2000, month=mm, day=dd)


def _sanitize_sheet_title(title: str, existing: set[str]) -> str:
    t = re.sub(r'[:\\/*?\[\]]', "-", str(title)).strip()
    if not t:
        t = "Sheet"
    t = t[:31]
    base = t
    i = 2
    while t in existing:
        suffix = f"_{i}"
        t = (base[: 31 - len(suffix)] + suffix)
        i += 1
    return t


def _find_col(df: pd.DataFrame, keywords: list[str]) -> Optional[str]:
    cols = [str(c) for c in df.columns]
    for c in cols:
        if all(k in c for k in keywords):
            return c
    return None


def _copy_sheet(wb, base_ws, title):
    ws = wb.copy_worksheet(base_ws)
    ws.title = _sanitize_sheet_title(title, set(wb.sheetnames))
    # copy_worksheetでは印刷範囲・タイトル・改ページが複製されない。
    if base_ws.print_area:
        ws.print_area = str(base_ws.print_area)
    ws.print_title_rows = base_ws.print_title_rows
    ws.print_title_cols = base_ws.print_title_cols
    ws.row_breaks = copy(base_ws.row_breaks)
    ws.col_breaks = copy(base_ws.col_breaks)
    return ws


def _clear_detail_rows(ws):
    for r in range(DETAIL_START_ROW, DETAIL_END_ROW + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = None


def _format_delivery_date(value: str) -> str:
    """
    '4/18土' -> '4月18日'
    '4/18'   -> '4月18日'
    """
    if value is None:
        return ""

    s = str(value)
    m = re.search(r"(\d{1,2})/(\d{1,2})", s)
    if not m:
        return s

    mm = int(m.group(1))
    dd = int(m.group(2))
    return f"{mm}月{dd}日"


def _write_delivery_date(ws, delivery_value: str, is_tokuyou: bool):
    cell = TOKUYOU_DELIVERY_CELL if is_tokuyou else YUHOUSE_DELIVERY_CELL
    formatted = _format_delivery_date(delivery_value)
    ws[cell] = f"{formatted}納品分"


def _decimal_quantity(qty) -> Decimal:
    """数量を十進数にし、floatの末尾に現れた微小な計算誤差だけを除く。

    固定小数桁への丸めは行わない。floatに限り、4 ULP以内で有効桁が
    6桁以上短くなる場合だけ短い表現を採用する。通常の小数や末尾まで
    有効な精密値、Decimal・文字列で渡された数量はそのまま保持する。
    """
    try:
        value = Decimal(str(qty))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal(0)
    if not value.is_finite():
        return Decimal(0)

    if isinstance(qty, float) and value:
        digits = len(value.as_tuple().digits)
        # Excelへの保存・再読込時の十進変換によるずれも含めた微小な範囲。
        tolerance = Decimal.from_float(math.ulp(qty)) * 4
        for precision in range(1, digits - 5):
            candidate = Decimal(format(qty, f".{precision}g"))
            if abs(candidate - value) <= tolerance:
                return candidate
    return value


def _format_qty_with_unit(qty, unit) -> str:
    q = _decimal_quantity(qty)
    if not q:
        return ""
    q_str = format(q, "f")
    if "." in q_str:
        q_str = q_str.rstrip("0").rstrip(".")
    unit_str = "" if unit is None else str(unit).strip()
    return f"{q_str}{unit_str}"


def _write_row_tokuyou(ws, row_no: int, use_date: str, food_name: str, qty_res, qty_staff, unit):
    qty_res = _decimal_quantity(qty_res)
    qty_staff = _decimal_quantity(qty_staff)
    ws.cell(row_no, 1).value = use_date
    ws.cell(row_no, 2).value = food_name
    ws.cell(row_no, 4).value = qty_res if qty_res != 0 else None
    ws.cell(row_no, 5).value = qty_staff if qty_staff != 0 else None

    total = (qty_res or 0) + (qty_staff or 0)
    ws.cell(row_no, 6).value = _format_qty_with_unit(total, unit)


def _write_row_yuhouse(ws, row_no: int, use_date: str, food_name: str, qty_res, unit):
    ws.cell(row_no, 1).value = use_date
    ws.cell(row_no, 2).value = food_name
    ws.cell(row_no, 4).value = _format_qty_with_unit(qty_res, unit)


def generate_hokubu_order_workbook(
    kenshu_xlsx_path: str | Path,
    template_xlsm_path: str | Path,
    facility_mode: str,
    out_path: str | Path,
) -> Path:

    # ------------------------------------------------------------
    # 加工済み検収簿を読み込む
    # ------------------------------------------------------------
    df = pd.read_excel(
        kenshu_xlsx_path
    )

    # ------------------------------------------------------------
    # 必須基本列チェック
    # ------------------------------------------------------------
    required_cols = [
        "仕入先",
        "納品日",
        "使用日",
        "食品名",
        "単位",
    ]

    missing_basic_cols = [
        c
        for c in required_cols
        if c not in df.columns
    ]

    if missing_basic_cols:
        raise KeyError(
            "検収簿に必要な列が見つかりません："
            + "、".join(missing_basic_cols)
        )

    # ------------------------------------------------------------
    # 北部市場販売だけ抽出
    # ------------------------------------------------------------
    df["仕入先"] = (
        df["仕入先"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df = df[
        df["仕入先"] == SUPPLIER_NAME
    ].copy()

    if df.empty:
        raise ValueError(
            "北部市場販売のデータが見つかりません。"
        )

    # ------------------------------------------------------------
    # ①検収簿整形後の統一列名
    # ------------------------------------------------------------
    tok_res = "特養入所者"
    tok_staff = "特養職員"
    yuhouse_res = "ユーハウス"

    # ------------------------------------------------------------
    # 特養
    # ------------------------------------------------------------
    if facility_mode == "tokuyou":

        missing_cols = []

        if tok_res not in df.columns:
            missing_cols.append(
                tok_res
            )

        if tok_staff not in df.columns:
            missing_cols.append(
                tok_staff
            )

        if missing_cols:
            raise KeyError(
                "北部市場・特養発注書に必要な列が見つかりません："
                + "、".join(missing_cols)
                + "。①検収簿整形で作成した最新の加工済み検収簿を使用してください。"
            )

        qty_cols = [
            tok_res,
            tok_staff,
        ]

        sheet_name = TOKUYOU_SHEET
        is_tokuyou = True

    # ------------------------------------------------------------
    # ユーハウス
    # ------------------------------------------------------------
    elif facility_mode == "yuhouse":

        if yuhouse_res not in df.columns:
            raise KeyError(
                "『ユーハウス』列が見つかりません。"
                "①検収簿整形で作成した最新の加工済み検収簿を使用してください。"
            )

        qty_cols = [
            yuhouse_res,
        ]

        sheet_name = YUHOUSE_SHEET
        is_tokuyou = False

    else:
        raise ValueError(
            "facility_mode must be 'tokuyou' or 'yuhouse'"
        )

    # ------------------------------------------------------------
    # 数量を数値へ変換
    # ------------------------------------------------------------
    for col in qty_cols:
        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0).map(_decimal_quantity)

    # ------------------------------------------------------------
    # 数量0の行を除外
    # ------------------------------------------------------------
    if is_tokuyou:

        # 入所者または職員のどちらかに数量があれば残す
        df = df[
            (df[tok_res] != 0)
            |
            (df[tok_staff] != 0)
        ].copy()

    else:

        df = df[
            df[yuhouse_res] != 0
        ].copy()

    if df.empty:
        raise ValueError(
            "北部市場販売の注文数量が入力されているデータがありません。"
        )

    # ------------------------------------------------------------
    # 同一食品をまとめる
    # ------------------------------------------------------------
    group_cols = [
        "納品日",
        "使用日",
        "食品名",
        "単位",
    ]

    grouped = (
        df.groupby(
            group_cols,
            dropna=False
        )[qty_cols]
        .sum()
        .reset_index()
    )

    # ------------------------------------------------------------
    # 日付順に並べる
    # ------------------------------------------------------------
    grouped["納品日_dt"] = (
        grouped["納品日"]
        .apply(_parse_mmdd)
    )

    grouped["使用日_dt"] = (
        grouped["使用日"]
        .apply(_parse_mmdd)
    )

    grouped = grouped.sort_values(
        [
            "納品日_dt",
            "使用日_dt",
            "食品名",
        ],
        na_position="last"
    ).reset_index(
        drop=True
    )

    # ------------------------------------------------------------
    # テンプレートを開く
    # ------------------------------------------------------------
    wb = openpyxl.load_workbook(
        template_xlsm_path,
        keep_vba=True
    )

    if sheet_name not in wb.sheetnames:
        raise KeyError(
            f"テンプレートに『{sheet_name}』シートが見つかりません。"
        )

    base_ws = wb[
        sheet_name
    ]

    current_delivery = None
    current_use_date = None

    page_no = 0
    row_in_page = 0
    ws = None
    output_sheets = []

    # ------------------------------------------------------------
    # 発注書作成
    # ------------------------------------------------------------
    for _, rec in grouped.iterrows():

        delivery = str(
            rec["納品日"]
        )

        use_date = str(
            rec["使用日"]
        )

        food_name = str(
            rec["食品名"]
        )

        unit = rec["単位"]

        # --------------------------------------------------------
        # 納品日・使用日が変わったら新しいシート
        # または12行を超えたら新しいページ
        # --------------------------------------------------------
        if (
            current_delivery != delivery
            or current_use_date != use_date
            or row_in_page >= ROWS_PER_PAGE
        ):

            current_delivery = delivery
            current_use_date = use_date

            page_no += 1

            facility_text = (
                "特養"
                if is_tokuyou
                else "ユーハウス"
            )

            title = (
                f"{delivery}_"
                f"{use_date}_"
                f"{facility_text}_"
                f"{page_no}"
            )

            ws = _copy_sheet(
                wb,
                base_ws,
                title
            )
            output_sheets.append(ws)

            _clear_detail_rows(
                ws
            )

            _write_delivery_date(
                ws,
                delivery,
                is_tokuyou=is_tokuyou
            )

            row_in_page = 0

        target_row = (
            DETAIL_START_ROW
            + row_in_page
        )

        # --------------------------------------------------------
        # 特養
        # --------------------------------------------------------
        if is_tokuyou:

            _write_row_tokuyou(
                ws=ws,
                row_no=target_row,
                use_date=use_date,
                food_name=food_name,
                qty_res=rec[tok_res],
                qty_staff=rec[tok_staff],
                unit=unit,
            )

        # --------------------------------------------------------
        # ユーハウス
        # --------------------------------------------------------
        else:

            _write_row_yuhouse(
                ws=ws,
                row_no=target_row,
                use_date=use_date,
                food_name=food_name,
                qty_res=rec[yuhouse_res],
                unit=unit,
            )

        row_in_page += 1

    # ------------------------------------------------------------
    # 作成済みの各シートに共通のフォント設定を適用
    # ------------------------------------------------------------
    for output_ws in output_sheets:
        _standardize_order_fonts(output_ws, is_tokuyou=is_tokuyou)

    # ------------------------------------------------------------
    # 元テンプレートシート削除
    # ------------------------------------------------------------
    if base_ws.title in wb.sheetnames:
        wb.remove(
            base_ws
        )

    # ------------------------------------------------------------
    # 保存
    # ------------------------------------------------------------
    out_path = Path(
        out_path
    )

    out_path.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    wb.save(
        out_path
    )
    wb.close()
    if wb.vba_archive is not None:
        wb.vba_archive.close()

    return out_path

def generate_hokubu_order_forms_both_facilities(
    kenshu_xlsx_path: str | Path,
    template_xlsm_path: str | Path,
    out_dir: str | Path,
    out_prefix: str = "北部市場発注書",
):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    tokuyou_path = out_dir / f"{out_prefix}_特養.xlsm"
    yuhouse_path = out_dir / f"{out_prefix}_ユーハウス.xlsm"

    p1 = generate_hokubu_order_workbook(
        kenshu_xlsx_path=kenshu_xlsx_path,
        template_xlsm_path=template_xlsm_path,
        facility_mode="tokuyou",
        out_path=tokuyou_path,
    )
    p2 = generate_hokubu_order_workbook(
        kenshu_xlsx_path=kenshu_xlsx_path,
        template_xlsm_path=template_xlsm_path,
        facility_mode="yuhouse",
        out_path=yuhouse_path,
    )
    return p1, p2

