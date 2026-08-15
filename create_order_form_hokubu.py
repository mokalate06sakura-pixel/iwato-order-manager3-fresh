

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
