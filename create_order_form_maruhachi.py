import re
from pathlib import Path
from typing import Dict, Tuple, List, Optional

import pandas as pd
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

def find_col_by_keywords(df: pd.DataFrame, keywords: list[str]) -> str:
    """
    df.columns から、keywords をすべて含む列名を1つ返す。
    見つからない場合は KeyError を投げる（列一覧も出す）。
    """
    cols = [str(c) for c in df.columns]
    for c in cols:
        if all(k in c for k in keywords):
            return c
    raise KeyError(f"列が見つかりません: keywords={keywords}\n利用可能な列:\n" + "\n".join(cols))

INVALID_SHEET_CHARS = r'[:\\/*?\[\]]'  # Excelで禁止される文字

def sanitize_sheet_title(title: str, existing: set[str]) -> str:
    """
    Excelのシート名制約に合わせて安全化：
    - 禁止文字を置換
    - 先頭/末尾の ' を避ける
    - 最大31文字
    - 重複する場合は _2, _3... を付与
    """
    t = str(title)
    t = re.sub(INVALID_SHEET_CHARS, "-", t)  # / を含む禁止文字を - に
    t = t.strip()
    if t.startswith("'"):
        t = t[1:]
    if t.endswith("'"):
        t = t[:-1]
    if not t:
        t = "Sheet"

    # 31文字制限
    t = t[:31]

    # 重複回避（_2, _3…）
    base = t
    i = 2
    while t in existing:
        suffix = f"_{i}"
        t = (base[: 31 - len(suffix)] + suffix)
        i += 1
    return t



# ====== 固定（今回のサンプルに合わせた列名） ======
COL_SUPPLIER = "仕入先"
COL_USE_DATE = "使用日"
COL_FOOD_NAME = "食品名"
COL_SPEC = "換算値"

SUPPLIER_NAME = "丸八ヒロタ"

# 施設別 数量列（固定しない：find_col_by_keywordsで自動検出）
# COL_TOKUYOU_RESIDENT = ...
# COL_TOKUYOU_STAFF = ...
# COL_YUHOUSE_RESIDENT = ...

# ====== テンプレ仕様（ユーザー確定事項） ======
TEMPLATE_SHEET_NAME = "丸八ヒロタ発注書"
TAG_SHEET_NAME = "タグ"

# 固定品目欄：ヘッダが5行目、明細が6行目～（コード列Bの連続で判定）
FIXED_HEADER_ROW = 5
FIXED_FIRST_ROW = 6

# 追記欄
APPEND_START_ROW = 24
APPEND_MAX_ROWS = 7  # 24～30

# 数量の書き込み列（テンプレ）
COL_OUT_USE_DATE = 1   # A列
COL_OUT_CODE     = 2   # B列（追記欄のコードは基本空）
COL_OUT_NAME_1   = 3   # C列（テンプレ上「☆」がある列）
COL_OUT_NAME_2   = 4   # D列（実品名が入っている列）
COL_OUT_SPEC     = 5   # E列
COL_OUT_RESIDENT = 6   # F列
COL_OUT_STAFF    = 7   # G列


def _norm(s: object) -> str:
    """品名照合用の正規化（表記ゆれの最低限対策）"""
    if s is None:
        return ""
    t = str(s)
    t = t.replace("\u3000", " ")  # 全角スペース→半角
    t = re.sub(r"\s+", " ", t)    # 連続空白/改行/タブ整理
    t = t.strip()
    return t


def load_tag_mapping(tag_xlsm_path: str | Path) -> Dict[str, Tuple[str, str, str]]:
    """
    タグシートから辞書を作る：
      key: ハートミール食品名（社内名）正規化
      val: (商品コード, 丸八品名, 規格)
    """
    tag_xlsm_path = Path(tag_xlsm_path)
    wb = openpyxl.load_workbook(tag_xlsm_path, data_only=True, keep_vba=True)
    ws = wb[TAG_SHEET_NAME]

    # ヘッダは1行目固定（今回のファイル構造に合わせる）
    # A:商品コード B:品名 C:規格 D:ハートミール食品名
    mapping: Dict[str, Tuple[str, str, str]] = {}
    for r in range(2, ws.max_row + 1):
        code = ws.cell(r, 1).value
        maru_name = ws.cell(r, 2).value
        spec = ws.cell(r, 3).value
        heart_name = ws.cell(r, 4).value

        k = _norm(heart_name)
        if not k:
            continue
        if code is None:
            continue

        mapping[k] = (str(code), str(maru_name or ""), str(spec or ""))

    return mapping


def _read_kenshu(kenshu_xlsx_path: str | Path) -> pd.DataFrame:
    df = pd.read_excel(Path(kenshu_xlsx_path))
    # 仕入先が丸八のみ
    df = df[df[COL_SUPPLIER].astype(str) == SUPPLIER_NAME].copy()
    # 使用日が空の行は除外
    df = df[df[COL_USE_DATE].notna()].copy()
    return df


def _build_fixed_row_index(ws: Worksheet) -> Dict[str, int]:
    """
    テンプレの固定品目欄（コード列B）から
      code -> row_number
    を作る
    """
    idx: Dict[str, int] = {}
    r = FIXED_FIRST_ROW
    while r < APPEND_START_ROW:
        code = ws.cell(r, COL_OUT_CODE).value
        if code is None or str(code).strip() == "":
            break
        idx[str(code).strip()] = r
        r += 1
    return idx


def _clear_sheet_quantities(ws: Worksheet) -> None:
    """数量・使用日・追記欄をクリア（固定品目行は残す）"""
    # 固定品目欄：A(使用日) F(入所者) G(職員) をクリア
    r = FIXED_FIRST_ROW
    while r < APPEND_START_ROW:
        code = ws.cell(r, COL_OUT_CODE).value
        if code is None or str(code).strip() == "":
            break
        ws.cell(r, COL_OUT_USE_DATE).value = None
        ws.cell(r, COL_OUT_RESIDENT).value = None
        ws.cell(r, COL_OUT_STAFF).value = None
        r += 1

    # 追記欄：24～30行の A,B,D,E,F,G をクリア（Cは☆列なので触らない）
    for rr in range(APPEND_START_ROW, APPEND_START_ROW + APPEND_MAX_ROWS):
        ws.cell(rr, COL_OUT_USE_DATE).value = None
        ws.cell(rr, COL_OUT_CODE).value = None
        ws.cell(rr, COL_OUT_NAME_2).value = None
        ws.cell(rr, COL_OUT_SPEC).value = None
        ws.cell(rr, COL_OUT_RESIDENT).value = None
        ws.cell(rr, COL_OUT_STAFF).value = None


def _write_append_row(
    ws: Worksheet,
    rr: int,
    use_date: object,
    food_name: str,
    spec: str,
    qty_resident: float,
    qty_staff: float,
) -> None:
    ws.cell(rr, COL_OUT_USE_DATE).value = use_date
    # コードは不明なので空欄のまま（要件）
    ws.cell(rr, COL_OUT_NAME_2).value = food_name
    ws.cell(rr, COL_OUT_SPEC).value = spec
    ws.cell(rr, COL_OUT_RESIDENT).value = qty_resident if qty_resident != 0 else None
    ws.cell(rr, COL_OUT_STAFF).value = qty_staff if qty_staff != 0 else None


def _copy_base_sheet(wb, base_ws, title):
    ws2 = wb.copy_worksheet(base_ws)

    existing = set(wb.sheetnames)
    safe_title = sanitize_sheet_title(title, existing)
    ws2.title = safe_title

    _clear_sheet_quantities(ws2)
    return ws2


def generate_maruhachi_order_workbook(
    kenshu_xlsx_path: str | Path,
    template_xlsm_path: str | Path,
    tag_xlsm_path: str | Path,
    facility_mode: str,
    out_path: str | Path,
) -> Path:
    """
    facility_mode:
      - "tokuyou" : 特養（入所者=I列, 職員=J列相当）
      - "yuhouse" : ユーハウス（入所者=K列相当, 職員は空）
    """
    if facility_mode not in ("tokuyou", "yuhouse"):
        raise ValueError("facility_mode must be 'tokuyou' or 'yuhouse'")

    kenshu_xlsx_path = Path(kenshu_xlsx_path)
    template_xlsm_path = Path(template_xlsm_path)
    tag_xlsm_path = Path(tag_xlsm_path)
    out_path = Path(out_path)

    # データ読み込み
    df = _read_kenshu(kenshu_xlsx_path)

    # 施設別に数量列を選択（列名ゆれ対策で自動検出）
    if facility_mode == "tokuyou":
        col_res = find_col_by_keywords(df, ["介護老人福祉施設いわと", "入所者"])
        col_staff = find_col_by_keywords(df, ["介護老人福祉施設いわと", "職員"])
    else:
        # ユーハウス（職員列は使わない運用）
        try:
            col_res = find_col_by_keywords(df, ["ケアハウス", "入所者"])
        except KeyError:
            col_res = find_col_by_keywords(df, ["ユーハウス", "入所者"])
        col_staff = None

    
    # タグ辞書（社内名→(コード,丸八品名,規格)）
    tag_map = load_tag_mapping(tag_xlsm_path)

    # テンプレWB（xlsm）
    wb = openpyxl.load_workbook(template_xlsm_path, keep_vba=True)
    base_ws = wb[TEMPLATE_SHEET_NAME]
    fixed_row_index = _build_fixed_row_index(base_ws)

    # baseをクリア（残してコピー元にする）
    _clear_sheet_quantities(base_ws)

    # 使用日ごとに処理
    # サンプルは "3/23月" のような文字列なので、そのまま使う（Excel表示も自然）
    use_dates = sorted(df[COL_USE_DATE].dropna().astype(str).unique().tolist())

    created_sheets = []
    for use_date in use_dates:
        ddf = df[df[COL_USE_DATE].astype(str) == use_date].copy()

        # 品目ごとに合計
        ddf[col_res] = pd.to_numeric(ddf[col_res], errors="coerce").fillna(0)

        if col_staff is not None:
            ddf[col_staff] = pd.to_numeric(ddf[col_staff], errors="coerce").fillna(0)
            col_staff_tmp = col_staff
        else:
            ddf["_staff"] = 0
            col_staff_tmp = "_staff"

        grouped = (
            ddf.groupby([COL_FOOD_NAME, COL_SPEC], dropna=False)[[col_res, col_staff_tmp]]
            .sum()
            .reset_index()
        )
        # シート作成（1ページ目）
        sheet_title = str(use_date)
        ws = _copy_base_sheet(wb, base_ws, sheet_title)
        created_sheets.append(ws.title)

        # 追記管理
        append_items: List[Tuple[str, str, float, float]] = []

        # 固定欄へ転記 or 追記へ
        for _, row in grouped.iterrows():
            food_name = _norm(row[COL_FOOD_NAME])
            spec = str(row[COL_SPEC] or "")
            qty_res = float(row[col_res] or 0)
            qty_staff = float(row[col_staff_tmp] or 0)

            if qty_res == 0 and qty_staff == 0:
                continue

            tag_key = _norm(food_name)
            if tag_key in tag_map:
                code, maru_name, maru_spec = tag_map[tag_key]
                # テンプレ固定欄にコードがあるか
                if code in fixed_row_index:
                    rr = fixed_row_index[code]
                    # 注文がある行だけ使用日・数量を書く（空欄行は残す）
                    ws.cell(rr, COL_OUT_USE_DATE).value = use_date
                    ws.cell(rr, COL_OUT_RESIDENT).value = qty_res if qty_res != 0 else None
                    ws.cell(rr, COL_OUT_STAFF).value = qty_staff if qty_staff != 0 else None
                else:
                    # コードはあるが固定欄にない → 追記扱い（安全）
                    append_items.append((food_name, spec, qty_res, qty_staff))
            else:
                # コード紐づけ無し → 追記
                append_items.append((food_name, spec, qty_res, qty_staff))

        # 追記欄に書く（溢れたら次ページ）
        if append_items:
            page = 1
            pos = 0
            while pos < len(append_items):
                if page == 1:
                    cur_ws = ws
                else:
                    # 次ページ作成（同じ日付名 + _{page}ページ目）
                    cur_ws = _copy_base_sheet(wb, base_ws, f"{use_date}_{page}ページ目")
                    created_sheets.append(cur_ws.title)

                start = APPEND_START_ROW
                end = APPEND_START_ROW + APPEND_MAX_ROWS  # exclusive

                for rr in range(start, end):
                    if pos >= len(append_items):
                        break
                    f, sp, qr, qs = append_items[pos]
                    _write_append_row(cur_ws, rr, use_date, f, sp, qr, qs)
                    pos += 1

                page += 1

    # コピー元のベースシートは削除（出力をスッキリ）
    # ※ base_ws がそのまま残って良い運用なら、削除しないでもOK
    if base_ws.title in wb.sheetnames and base_ws.title not in created_sheets:
        pass
    else:
        # baseがコピー元として残っているので消す
        try:
            wb.remove(base_ws)
        except Exception:
            # 保護されている等の場合は残す
            pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def generate_maruhachi_order_forms_both_facilities(
    kenshu_xlsx_path: str | Path,
    template_xlsm_path: str | Path,
    tag_xlsm_path: str | Path,
    out_dir: str | Path,
    out_prefix: str = "丸八発注書",
) -> Tuple[Path, Path]:
    """
    特養とユーハウスを別ブックで出力
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    tokuyou_path = out_dir / f"{out_prefix}_特養.xlsm"
    yuhouse_path = out_dir / f"{out_prefix}_ユーハウス.xlsm"

    p1 = generate_maruhachi_order_workbook(
        kenshu_xlsx_path=kenshu_xlsx_path,
        template_xlsm_path=template_xlsm_path,
        tag_xlsm_path=tag_xlsm_path,
        facility_mode="tokuyou",
        out_path=tokuyou_path,
    )
    p2 = generate_maruhachi_order_workbook(
        kenshu_xlsx_path=kenshu_xlsx_path,
        template_xlsm_path=template_xlsm_path,
        tag_xlsm_path=tag_xlsm_path,
        facility_mode="yuhouse",
        out_path=yuhouse_path,
    )
    return p1, p2
