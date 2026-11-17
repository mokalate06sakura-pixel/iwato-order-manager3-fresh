"""SplitMenuDaily_A3_V4 のWeb版。元ブックのXML・画像を保持して日別シートを追加する。"""
from io import BytesIO
import posixpath
import re
from xml.dom import minidom
from xml.parsers.expat import ExpatError
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from openpyxl.formula.translate import Translator
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter, range_boundaries

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PREFIX = "日別_"
MEALS = {"朝食", "昼食", "夕食"}
OUTPUT_LAYOUT_REVISION = "daily-menu-v6-20261117"


def _children(node, name=None):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE and (name is None or n.localName == name)]


def _first(node, name):
    return next(iter(_children(node, name)), None)


def _text(node):
    if node is None:
        return ""
    return "".join(n.data if n.nodeType in (n.TEXT_NODE, n.CDATA_SECTION_NODE) else _text(n) for n in node.childNodes)


def _element(doc, name, attrs=None):
    node = doc.createElementNS(MAIN, name)
    for key, value in (attrs or {}).items():
        node.setAttribute(key, str(value))
    return node


def _xml(node):
    return node.toxml(encoding="utf-8")


def _target(base, target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target)) if not target.startswith("/") else target.lstrip("/")


class _Package:
    def __init__(self, data):
        try:
            with ZipFile(BytesIO(data)) as archive:
                self.files = {n: archive.read(n) for n in archive.namelist()}
            self.workbook = minidom.parseString(self.files["xl/workbook.xml"])
            self.rels = minidom.parseString(self.files["xl/_rels/workbook.xml.rels"])
            self.types = minidom.parseString(self.files["[Content_Types].xml"])
        except (BadZipFile, KeyError, ValueError, TypeError, ExpatError) as exc:
            raise ValueError("Excelブック（xlsx）を読み込めません。ファイルを確認してください。") from exc
        self.sheets = _first(self.workbook.documentElement, "sheets")
        self.shared = []
        if "xl/sharedStrings.xml" in self.files:
            root = minidom.parseString(self.files["xl/sharedStrings.xml"]).documentElement
            self.shared = ["".join(_text(t) for t in si.getElementsByTagNameNS(MAIN, "t")) for si in _children(root, "si")]

    def sheet_path(self, sheet):
        rid = sheet.getAttributeNS(REL, "id")
        rel = next(r for r in _children(self.rels.documentElement) if r.getAttribute("Id") == rid)
        return _target("xl/workbook.xml", rel.getAttribute("Target"))

    def cell_text(self, cell):
        kind = cell.getAttribute("t")
        if kind == "e":
            return ""
        if kind == "inlineStr":
            return "".join(_text(t) for t in cell.getElementsByTagNameNS(MAIN, "t"))
        value = _text(_first(cell, "v"))
        return self.shared[int(value)] if kind == "s" and value else value

    def source(self, name):
        sheet = next((s for s in _children(self.sheets) if s.getAttribute("name") == name), None)
        if sheet is None or name.startswith(PREFIX):
            raise ValueError("日別_で始まらない元の献立シートを選んでください。")
        path = self.sheet_path(sheet)
        return path, minidom.parseString(self.files[path])

    def remove_daily(self):
        sheets = _children(self.sheets)
        remaining = [s for s in sheets if not s.getAttribute("name").startswith(PREFIX)]
        if not remaining:
            raise ValueError("元のシートがありません。日別シートだけのブックは処理できません。")
        removed = [s for s in sheets if s not in remaining]
        removed_ids = {s.getAttributeNS(REL, "id") for s in removed}
        for s in removed:
            path = self.sheet_path(s)
            self.files.pop(path, None)
            self.files.pop(posixpath.join(posixpath.dirname(path), "_rels", posixpath.basename(path) + ".rels"), None)
            self.sheets.removeChild(s)
            for override in list(_children(self.types.documentElement)):
                if override.getAttribute("PartName") == "/" + path:
                    self.types.documentElement.removeChild(override)
        for rel in list(_children(self.rels.documentElement)):
            if rel.getAttribute("Id") in removed_ids:
                self.rels.documentElement.removeChild(rel)
        names = _first(self.workbook.documentElement, "definedNames")
        if names is not None:
            for name in list(_children(names)):
                if name.hasAttribute("localSheetId"):
                    old = int(name.getAttribute("localSheetId"))
                    if old >= len(sheets) or sheets[old] not in remaining:
                        names.removeChild(name)
                    else:
                        name.setAttribute("localSheetId", str(remaining.index(sheets[old])))
        return len(removed)

    def finish(self, active_index=0):
        # 古い計算チェーンは日別シートの作り直し後には使わない。
        for rel in list(_children(self.rels.documentElement)):
            if rel.getAttribute("Type").endswith("/calcChain"):
                self.files.pop(_target("xl/workbook.xml", rel.getAttribute("Target")), None)
                self.rels.documentElement.removeChild(rel)
        for override in list(_children(self.types.documentElement)):
            if override.getAttribute("ContentType").endswith("calcChain+xml"):
                self.types.documentElement.removeChild(override)
        for view in self.workbook.getElementsByTagNameNS(MAIN, "workbookView"):
            view.setAttribute("activeTab", str(active_index))
            view.setAttribute("firstSheet", "0")
        for index, sheet in enumerate(_children(self.sheets)):
            path = self.sheet_path(sheet)
            if not path.startswith("xl/worksheets/"):
                continue
            doc = minidom.parseString(self.files[path])
            changed = False
            for view in doc.getElementsByTagNameNS(MAIN, "sheetView"):
                selected = "1" if index == active_index else "0"
                if view.getAttribute("tabSelected") != selected:
                    view.setAttribute("tabSelected", selected)
                    changed = True
            if changed:
                self.files[path] = _xml(doc)
            doc.unlink()
        calc = _first(self.workbook.documentElement, "calcPr")
        if calc is not None:
            calc.setAttribute("fullCalcOnLoad", "1")
        self.files["xl/workbook.xml"] = _xml(self.workbook)
        self.files["xl/_rels/workbook.xml.rels"] = _xml(self.rels)
        self.files["[Content_Types].xml"] = _xml(self.types)
        result = BytesIO()
        with ZipFile(result, "w", ZIP_DEFLATED) as archive:
            for name, data in self.files.items():
                archive.writestr(name, data)
        return result.getvalue()


def list_menu_sheets(data):
    """画面で選べる元シートと、元ブックのアクティブシートを返す。"""
    book = _Package(data)
    sheets = _children(book.sheets)
    names = [s.getAttribute("name") for s in sheets if not s.getAttribute("name").startswith(PREFIX) and book.sheet_path(s).startswith("xl/worksheets/")]
    if not names:
        raise ValueError("処理できる元のワークシートがありません。")
    views = book.workbook.getElementsByTagNameNS(MAIN, "workbookView")
    active = int(views[0].getAttribute("activeTab") or 0) if views else 0
    selected = sheets[min(active, len(sheets) - 1)].getAttribute("name")
    return names, selected if selected in names else names[0]


def _find_stamps(book, source_path, source, date_rows, remove):
    count = 0
    groups = {row: [] for row in date_rows}
    rel_path = posixpath.join(posixpath.dirname(source_path), "_rels", posixpath.basename(source_path) + ".rels")
    if rel_path not in book.files:
        return count, groups
    rels = minidom.parseString(book.files[rel_path])
    for drawing in source.getElementsByTagNameNS(MAIN, "drawing"):
        rid = drawing.getAttributeNS(REL, "id")
        rel = next((r for r in _children(rels.documentElement) if r.getAttribute("Id") == rid), None)
        if rel is None:
            continue
        path = _target(source_path, rel.getAttribute("Target"))
        doc = minidom.parseString(book.files[path])
        changed = False
        for anchor in list(_children(doc.documentElement)):
            start = _first(anchor, "from")
            # VBAの実際の条件：TopLeftCell.Column >= 33（AG列以降）、日付行の3行上～当日。
            if start is None or _first(anchor, "pic") is None:
                continue
            col = int(_text(_first(start, "col")))
            row = int(_text(_first(start, "row"))) + 1
            if col >= 32 and any(d - 3 <= row <= d for d in date_rows):
                for date_row in date_rows:
                    if date_row - 3 <= row <= date_row:
                        stamp_doc = minidom.Document()
                        stamp_root = stamp_doc.importNode(doc.documentElement, False)
                        stamp_doc.appendChild(stamp_root)
                        stamp_root.appendChild(stamp_doc.importNode(anchor, True))
                        groups[date_row].append((path, _xml(stamp_doc)))
                        stamp_doc.unlink()
                if remove:
                    doc.documentElement.removeChild(anchor)
                    changed = True
                count += 1
        if changed:
            book.files[path] = _xml(doc)
        doc.unlink()
    return count, groups


def _add_stamps(book, sheet_path, sheet_id, date_row, stamps, right_align=False, source=None):
    """日付見出しのスタンプだけを、日別・Bシートの見出し位置にコピーする。"""
    if not stamps:
        return 0
    drawing_path = f"xl/drawings/daily{sheet_id}.xml"
    while drawing_path in book.files:
        drawing_path = drawing_path[:-4] + "x.xml"
    drawing_doc = None
    new_rels = minidom.Document()
    rel_root = new_rels.createElementNS(PKG, "Relationships")
    rel_root.setAttribute("xmlns", PKG)
    new_rels.appendChild(rel_root)
    rel_ids = {}
    offset = max(1, date_row - 2) - 1
    for original_path, xml in stamps:
        stamp_doc = minidom.parseString(xml)
        if drawing_doc is None:
            drawing_doc = minidom.Document()
            drawing_doc.appendChild(drawing_doc.importNode(stamp_doc.documentElement, False))
        anchor = drawing_doc.importNode(_children(stamp_doc.documentElement)[0], True)
        start = _first(anchor, "from")
        source_row = int(_text(_first(start, "row")))
        delta = max(0, source_row - offset) - source_row
        for tag in ("from", "to"):
            position = _first(anchor, tag)
            if position is not None:
                row_node = _first(position, "row")
                row_node.firstChild.data = str(max(0, int(_text(row_node)) + delta))
        if right_align:
            anchor = _right_aligned_stamp(book, drawing_doc, anchor, sheet_path, source, stamp_doc)
        original_rel_path = posixpath.join(posixpath.dirname(original_path), "_rels", posixpath.basename(original_path) + ".rels")
        original_rels = minidom.parseString(book.files[original_rel_path])
        for node in anchor.getElementsByTagName("*"):
            for attribute in ("embed", "link"):
                old_id = node.getAttributeNS(REL, attribute)
                if not old_id:
                    continue
                key = (original_path, old_id)
                if key not in rel_ids:
                    original_rel = next(r for r in _children(original_rels.documentElement) if r.getAttribute("Id") == old_id)
                    copied_rel = new_rels.importNode(original_rel, True)
                    new_id = f"rId{len(rel_ids) + 1}"
                    copied_rel.setAttribute("Id", new_id)
                    if original_rel.getAttribute("TargetMode") != "External":
                        copied_rel.setAttribute("Target", posixpath.relpath(_target(original_path, original_rel.getAttribute("Target")), posixpath.dirname(drawing_path)))
                    rel_root.appendChild(copied_rel)
                    rel_ids[key] = new_id
                node.setAttributeNS(REL, "r:" + attribute, rel_ids[key])
                node.setAttribute("xmlns:r", REL)
        drawing_doc.documentElement.appendChild(anchor)
        original_rels.unlink()
        stamp_doc.unlink()
    book.files[drawing_path] = _xml(drawing_doc)
    drawing_rels_path = posixpath.join(posixpath.dirname(drawing_path), "_rels", posixpath.basename(drawing_path) + ".rels")
    book.files[drawing_rels_path] = _xml(new_rels)
    drawing_doc.unlink()
    new_rels.unlink()
    worksheet = minidom.parseString(book.files[sheet_path])
    drawing = _element(worksheet, "drawing")
    drawing.setAttribute("xmlns:r", REL)
    drawing.setAttributeNS(REL, "r:id", "rIdStamps")
    worksheet.documentElement.appendChild(drawing)
    book.files[sheet_path] = _xml(worksheet)
    worksheet.unlink()
    rel_doc = minidom.Document()
    rel_root = rel_doc.createElementNS(PKG, "Relationships")
    rel_root.setAttribute("xmlns", PKG)
    rel_doc.appendChild(rel_root)
    relation = rel_doc.createElementNS(PKG, "Relationship")
    for key, value in {"Id": "rIdStamps", "Type": REL + "/drawing", "Target": posixpath.relpath(drawing_path, posixpath.dirname(sheet_path))}.items():
        relation.setAttribute(key, value)
    rel_root.appendChild(relation)
    sheet_rel_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
    book.files[sheet_rel_path] = _xml(rel_doc)
    rel_doc.unlink()
    override = book.types.createElementNS(CT, "Override")
    override.setAttribute("PartName", "/" + drawing_path)
    override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")
    book.types.documentElement.appendChild(override)
    return len(stamps)


def _column_specs(root, count=36):
    props = _first(root, "sheetFormatPr")
    default = props.getAttribute("defaultColWidth") if props is not None else ""
    result = [{"width": default or "9.140625"} for _ in range(count)]
    cols = _first(root, "cols")
    if cols is not None:
        for col in _children(cols):
            for index in range(max(1, int(col.getAttribute("min"))), min(count, int(col.getAttribute("max"))) + 1):
                result[index - 1].update({key: col.getAttribute(key) for key in col.attributes.keys() if key not in ("min", "max")})
    return result


def _digit_points(book):
    """Normal書体の文字幅をポイントへ換算する（実献立のMS Pゴシック11で実機照合）。"""
    styles = minidom.parseString(book.files["xl/styles.xml"])
    font_id = 0
    cell_styles = _first(styles.documentElement, "cellStyles")
    style_xfs = _first(styles.documentElement, "cellStyleXfs")
    if cell_styles is not None and style_xfs is not None:
        normal = next((s for s in _children(cell_styles) if s.getAttribute("builtinId") == "0"), None)
        if normal is not None:
            xf = _children(style_xfs)[int(normal.getAttribute("xfId") or 0)]
            font_id = int(xf.getAttribute("fontId") or 0)
    font = _children(_first(styles.documentElement, "fonts"))[font_id]
    size = _first(font, "sz")
    points = 5.4 * float(size.getAttribute("val") if size is not None else 11) / 11
    styles.unlink()
    return points


def _column_points(spec, digit_points):
    return 0 if spec.get("hidden") == "1" else float(spec["width"]) * digit_points


def _row_points(root):
    props = _first(root, "sheetFormatPr")
    default = float(props.getAttribute("defaultRowHeight") or 15) if props is not None else 15
    rows = _children(_first(root, "sheetData"), "row")
    return {int(r.getAttribute("r")): 0 if r.getAttribute("hidden") == "1" else float(r.getAttribute("ht") or default) for r in rows}, default


def _right_aligned_stamp(book, doc, anchor, sheet_path, source, stamp_doc):
    original = _children(stamp_doc.documentElement)[0]
    extent = _first(original, "ext")
    if extent is None:
        extent = next((n for n in original.getElementsByTagName("*") if n.localName == "ext" and n.hasAttribute("cx") and n.hasAttribute("cy")), None)
    digit = _digit_points(book)
    if extent is not None:
        width, height = int(extent.getAttribute("cx")), int(extent.getAttribute("cy"))
    else:
        widths = [_column_points(s, digit) for s in _column_specs(source.documentElement)]
        heights, default = _row_points(source.documentElement)
        def position(tag):
            marker = _first(original, tag)
            col, row = int(_text(_first(marker, "col"))), int(_text(_first(marker, "row")))
            x = sum(widths[:col]) * 12700 + int(_text(_first(marker, "colOff")))
            y = sum(heights.get(r, default) for r in range(1, row + 1)) * 12700 + int(_text(_first(marker, "rowOff")))
            return x, y
        left, top = position("from")
        right, bottom = position("to")
        width, height = max(1, round(right - left)), max(1, round(bottom - top))
    worksheet = minidom.parseString(book.files[sheet_path])
    widths = [_column_points(s, digit) for s in _column_specs(worksheet.documentElement, 27)]
    worksheet.unlink()
    left = max(0, sum(widths) - width / 12700)
    column, remainder = 0, left
    while column < 26 and remainder >= widths[column]:
        remainder -= widths[column]
        column += 1
    namespace = anchor.namespaceURI
    prefix = anchor.prefix + ":" if anchor.prefix else ""
    def node(name, attrs=None, text=None):
        element = doc.createElementNS(namespace, prefix + name)
        for key, value in (attrs or {}).items():
            element.setAttribute(key, str(value))
        if text is not None:
            element.appendChild(doc.createTextNode(str(text)))
        return element
    replacement = node("oneCellAnchor")
    marker = node("from")
    for name, value in (("col", column), ("colOff", round(remainder * 12700)), ("row", 0), ("rowOff", 0)):
        marker.appendChild(node(name, text=value))
    replacement.appendChild(marker)
    replacement.appendChild(node("ext", {"cx": width, "cy": height}))
    for name in ("pic", "clientData"):
        item = _first(anchor, name)
        if item is not None:
            replacement.appendChild(item.cloneNode(True))
    return replacement


def _delete_iwato_formula(expression):
    tokens = Tokenizer("=" + expression).items
    pattern = re.compile(r"(\$?)([A-Z]{1,3})(\$?\d*)$")
    for token in tokens:
        if token.subtype != "RANGE" or "!" in token.value:
            continue
        parts = token.value.split(":")
        matches = [pattern.fullmatch(part) for part in parts]
        if not all(matches) or (len(parts) == 1 and not re.search(r"\d", parts[0])):
            continue
        columns = [column_index_from_string(m[2]) for m in matches]
        if max(columns) <= 9:
            token.value = "#REF!"
        else:
            token.value = ":".join(m[1] + get_column_letter(max(10, col) - 9) + m[3] for m, col in zip(matches, columns))
    return "".join(t.value for t in tokens)


def _restore_label_style(book, source_style, destination_style):
    if source_style == destination_style:
        return source_style
    if not hasattr(book, "layout_styles"):
        book.layout_styles = minidom.parseString(book.files["xl/styles.xml"])
        book.label_styles = {}
    key = source_style, destination_style
    if key in book.label_styles:
        return book.label_styles[key]
    styles = book.layout_styles
    xfs = _first(styles.documentElement, "cellXfs")
    source, destination = _children(xfs)[source_style], _children(xfs)[destination_style]
    xf = destination.cloneNode(True)
    xf.setAttribute("fontId", source.getAttribute("fontId") or "0")
    xf.setAttribute("applyFont", "1")
    fill = _children(_first(styles.documentElement, "fills"))[int(source.getAttribute("fillId") or 0)]
    pattern = _first(fill, "patternFill")
    if pattern is not None and pattern.getAttribute("patternType") not in ("", "none"):
        xf.setAttribute("fillId", source.getAttribute("fillId"))
        xf.setAttribute("applyFill", "1")
    src_alignment, alignment = _first(source, "alignment"), _first(xf, "alignment")
    if alignment is None:
        alignment = _element(styles, "alignment")
        xf.appendChild(alignment)
    if src_alignment is not None and src_alignment.hasAttribute("horizontal"):
        alignment.setAttribute("horizontal", src_alignment.getAttribute("horizontal"))
    elif alignment.hasAttribute("horizontal"):
        alignment.removeAttribute("horizontal")
    xf.setAttribute("applyAlignment", "1")
    result = len(_children(xfs))
    xfs.appendChild(xf)
    xfs.setAttribute("count", str(result + 1))
    book.label_styles[key] = result
    book.files["xl/styles.xml"] = _xml(styles)
    return result


def _compact_layout(book, xml, equalize):
    doc = minidom.parseString(xml)
    root = doc.documentElement
    rows = _children(_first(root, "sheetData"), "row")
    for row in rows:
        saved = None
        for cell in _children(row, "c"):
            column, number = coordinate_from_string(cell.getAttribute("r"))
            index = column_index_from_string(column)
            if index == 1:
                text = book.cell_text(cell)
                if text.startswith(("現場", "年月日")) or text in MEALS:
                    saved = cell.cloneNode(True)
            if index <= 9:
                row.removeChild(cell)
            else:
                cell.setAttribute("r", f"{get_column_letter(index - 9)}{number}")
                formula = _first(cell, "f")
                if formula is not None:
                    expression = _delete_iwato_formula(_text(formula))
                    for child in list(formula.childNodes):
                        formula.removeChild(child)
                    formula.appendChild(doc.createTextNode(expression))
        if saved is not None:
            destination = next((c for c in _children(row, "c") if c.getAttribute("r") == saved.getAttribute("r")), None)
            old_style = int(destination.getAttribute("s") or 0) if destination is not None else 0
            saved.setAttribute("s", str(_restore_label_style(book, int(saved.getAttribute("s") or 0), old_style)))
            if destination is not None:
                row.replaceChild(saved, destination)
            else:
                row.insertBefore(saved, row.firstChild)
        if row.hasAttribute("spans"):
            row.setAttribute("spans", "1:27")
    merges = _first(root, "mergeCells")
    if merges is not None:
        for merge in list(_children(merges)):
            a, b, c, d = range_boundaries(merge.getAttribute("ref"))
            if c <= 9:
                merges.removeChild(merge)
            else:
                merge.setAttribute("ref", f"{get_column_letter(max(a, 10) - 9)}{b}:{get_column_letter(c - 9)}{d}")
        merges.setAttribute("count", str(len(_children(merges))))
    specs = _column_specs(root)[9:36]
    if equalize:
        margins = _first(root, "pageMargins")
        margin = lambda name, default: float(margins.getAttribute(name) or default) if margins is not None else default
        printable_width = 297 / 25.4 * 72 - 72 * (margin("left", 0.7) + margin("right", 0.7))
        printable_height = 420 / 25.4 * 72 - 72 * (margin("top", 0.75) + margin("bottom", 0.75))
        if min(printable_width, printable_height) <= 0:
            raise ValueError("印刷余白がA3用紙より大きいため、幅を調整できません。")
        heights, default = _row_points(root)
        content_height = sum(heights.values())
        target = printable_width * max(1, content_height / printable_height) / 3
        digit = _digit_points(book)
        for block in range(3):
            group = specs[block * 9:(block + 1) * 9]
            total = sum(_column_points(s, digit) for s in group)
            if total > 0:
                for spec in group:
                    width = float(spec["width"]) * target / total
                    if width > 255:
                        raise ValueError("行数が多く、A3幅に合わせる列幅がExcelの上限を超えます。日数や行数を確認してください。")
                    spec["width"] = str(round(width * 256) / 256)
                    spec["customWidth"] = "1"
    cols = _first(root, "cols")
    if cols is None:
        cols = _element(doc, "cols")
        root.insertBefore(cols, _first(root, "sheetData"))
    for child in list(cols.childNodes):
        cols.removeChild(child)
    for index, spec in enumerate(specs, 1):
        cols.appendChild(_element(doc, "col", {**spec, "min": index, "max": index}))
    dimension = _first(root, "dimension")
    if dimension is not None:
        dimension.setAttribute("ref", f"A1:AA{len(rows)}")
    result = _xml(doc)
    doc.unlink()
    return result


def _daily_xml(source, selected_rows, shared_formulas):
    doc = minidom.Document()
    root = doc.importNode(source.documentElement, False)
    doc.appendChild(root)
    keep = {"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "mergeCells", "printOptions", "pageMargins", "pageSetup", "headerFooter"}
    for node in _children(source.documentElement):
        if node.localName in keep:
            root.appendChild(doc.importNode(node, node.localName not in ("sheetData", "mergeCells")))
    mapping = {old: new for new, old in enumerate(selected_rows, 1)}
    original_data = _first(source.documentElement, "sheetData")
    rows = {int(r.getAttribute("r")): r for r in _children(original_data, "row")}
    data = _first(root, "sheetData")
    for node in list(data.childNodes):
        data.removeChild(node)
    for old, new in mapping.items():
        row = doc.importNode(rows[old], True) if old in rows else _element(doc, "row")
        row.setAttribute("r", str(new))
        for cell in list(_children(row, "c")):
            origin = cell.getAttribute("r")
            column, _ = coordinate_from_string(origin)
            if column_index_from_string(column) > 36:
                row.removeChild(cell)
                continue
            dest = f"{column}{new}"
            cell.setAttribute("r", dest)
            formula = _first(cell, "f")
            if formula is not None:
                if formula.getAttribute("t") not in ("", "normal", "shared"):
                    raise ValueError("配列数式のある献立には対応していません。通常のセル数式に変更してください。")
                expression = _text(formula)
                if formula.getAttribute("t") == "shared" and not expression:
                    master, expression = shared_formulas[formula.getAttribute("si")]
                    expression = Translator("=" + expression, origin=master).translate_formula(origin)[1:]
                translated = Translator("=" + expression, origin=origin).translate_formula(dest)[1:]
                for attr in list(formula.attributes.keys()):
                    formula.removeAttribute(attr)
                for node in list(formula.childNodes):
                    formula.removeChild(node)
                formula.appendChild(doc.createTextNode(translated))
        data.appendChild(row)
    dim = _first(root, "dimension")
    if dim is not None:
        dim.setAttribute("ref", f"A1:AJ{len(selected_rows)}")
    merges = _first(root, "mergeCells")
    if merges is not None:
        for original_merge in _children(_first(source.documentElement, "mergeCells")):
            a, b, c, d = range_boundaries(original_merge.getAttribute("ref"))
            hits = [mapping[r] for r in range(b, d + 1) if r in mapping]
            if hits and c <= 36:
                merge = doc.importNode(original_merge, True)
                left = re.match(r"[A-Z]+", merge.getAttribute("ref")).group()
                right = re.search(r":([A-Z]+)", merge.getAttribute("ref"))
                right = right.group(1) if right else left
                merge.setAttribute("ref", f"{left}{min(hits)}:{right}{max(hits)}")
                merges.appendChild(merge)
        merges.setAttribute("count", str(len(_children(merges))))
    for view in root.getElementsByTagNameNS(MAIN, "sheetView"):
        view.setAttribute("tabSelected", "0")
        view.setAttribute("topLeftCell", "A1")
        for child in list(_children(view)):
            view.removeChild(child)
    props = _first(root, "sheetPr")
    if props is None:
        props = _element(doc, "sheetPr")
        root.insertBefore(props, root.firstChild)
    if props.hasAttribute("codeName"):
        props.removeAttribute("codeName")
    fit = _first(props, "pageSetUpPr")
    if fit is None:
        fit = _element(doc, "pageSetUpPr")
        props.appendChild(fit)
    fit.setAttribute("fitToPage", "1")
    fit.setAttribute("autoPageBreaks", "0")
    setup = _first(root, "pageSetup")
    if setup is None:
        setup = _element(doc, "pageSetup")
        footer = _first(root, "headerFooter")
        root.insertBefore(setup, footer) if footer is not None else root.appendChild(setup)
    if setup.hasAttributeNS(REL, "id"):
        setup.removeAttributeNS(REL, "id")
    if setup.hasAttribute("scale"):
        setup.removeAttribute("scale")
    if setup.hasAttribute("usePrinterDefaults"):
        setup.removeAttribute("usePrinterDefaults")
    # usePrinterDefaults=Falseを明示するとExcel実機でA4へ戻るため、VBA出力同様に省略する。
    for key, value in {"paperSize": "8", "orientation": "portrait", "fitToWidth": "1", "fitToHeight": "1"}.items():
        setup.setAttribute(key, value)
    result = _xml(doc)
    doc.unlink()
    return result


def split_menu_daily(data, sheet_name, include_stamps=False, remove_iwato=False, equalize_widths=False):
    """元シートを残し、日別A3縦1ページのブックと処理件数を返す。"""
    book = _Package(data)
    if equalize_widths and not remove_iwato:
        raise ValueError("3区分の幅を調整する場合は「いわと常食を除く」も選択してください。")
    source_path, source = book.source(sheet_name)
    rows = _children(_first(source.documentElement, "sheetData"), "row")
    texts = {int(row.getAttribute("r")): {coordinate_from_string(c.getAttribute("r"))[0]: book.cell_text(c) for c in _children(row, "c")} for row in rows}
    dates = [(r, cols["A"]) for r, cols in texts.items() if cols.get("A", "").startswith("年月日：")]
    dates.sort()
    if not dates:
        raise ValueError("A列に「年月日：」で始まる日付行がありません。元の献立シートを選んでください。")
    last = max((r for r, cols in texts.items() if any(cols.values())), default=0)
    plans = []
    plan_dates = {}
    b_count = 0
    for index, (date_row, label) in enumerate(dates):
        match = re.search(r"年\s*(\d+)月\s*(\d+)日", label[4:])
        if match is None:
            raise ValueError(f"{date_row}行目の日付を読み取れません。例：年月日：2026年11月17日")
        name = PREFIX + f"{int(match[1])}月{int(match[2])}日"
        start = max(1, date_row - 2)
        end = dates[index + 1][0] - 3 if index + 1 < len(dates) else last
        markers = [r for r in range(start, end + 1) if "<B>" in texts.get(r, {}).values()]
        if len(markers) > 1:
            raise ValueError(f"{name}に<B>が複数あります。1日1か所のセレクト献立に対応しています。")
        main_rows = list(range(start, end + 1))
        b_rows = None
        if markers:
            b_start = markers[0]
            next_meal = next((r for r in range(b_start + 1, end + 1) if texts.get(r, {}).get("A") in MEALS), end + 1)
            meal = next((r for r in range(b_start - 1, start - 1, -1) if texts.get(r, {}).get("A") in MEALS), None)
            b_rows = list(range(start, start + 4)) + ([meal] if meal else []) + list(range(b_start, next_meal))
            if len(b_rows) != len(set(b_rows)):
                raise ValueError(f"{name}のセレクト献立の位置を確認してください。先頭4行とBの帯が重複しています。")
            main_rows = [r for r in main_rows if not b_start <= r < next_meal]
            b_count += 1
        plans.append((name, main_rows))
        plan_dates[name] = date_row
        if b_rows:
            plans.append((name + "_B", b_rows))
            plan_dates[name + "_B"] = date_row
    if len({name for name, _ in plans}) != len(plans):
        raise ValueError("同じ月日の行が複数あります。1か月分の献立を選んでください。")
    shared = {}
    for cell in source.getElementsByTagNameNS(MAIN, "c"):
        formula = _first(cell, "f")
        if formula is not None and formula.getAttribute("t") == "shared" and _text(formula):
            shared[formula.getAttribute("si")] = (cell.getAttribute("r"), _text(formula))
    removed = book.remove_daily()
    stamps, stamp_groups = _find_stamps(book, source_path, source, [r for r, _ in dates], remove=not include_stamps and not remove_iwato)
    stamp_copies = 0
    first_index = len(_children(book.sheets))
    sheet_id = max(int(s.getAttribute("sheetId")) for s in _children(book.sheets))
    names = _first(book.workbook.documentElement, "definedNames")
    if names is None:
        names = _element(book.workbook, "definedNames")
        calc = _first(book.workbook.documentElement, "calcPr")
        book.workbook.documentElement.insertBefore(names, calc) if calc is not None else book.workbook.documentElement.appendChild(names)
    for offset, (name, selected_rows) in enumerate(plans):
        sheet_id += 1
        path = f"xl/worksheets/daily{sheet_id}.xml"
        while path in book.files:
            sheet_id += 1
            path = f"xl/worksheets/daily{sheet_id}.xml"
        rid = f"rIdDaily{sheet_id}"
        while any(r.getAttribute("Id") == rid for r in _children(book.rels.documentElement)):
            rid += "x"
        sheet = _element(book.workbook, "sheet", {"name": name, "sheetId": sheet_id})
        sheet.setAttribute("xmlns:r", REL)
        sheet.setAttributeNS(REL, "r:id", rid)
        book.sheets.appendChild(sheet)
        rel = book.rels.createElementNS(PKG, "Relationship")
        for key, val in {"Id": rid, "Type": REL + "/worksheet", "Target": path[3:]}.items():
            rel.setAttribute(key, val)
        book.rels.documentElement.appendChild(rel)
        override = book.types.createElementNS(CT, "Override")
        override.setAttribute("PartName", "/" + path)
        override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        book.types.documentElement.appendChild(override)
        area = _element(book.workbook, "definedName", {"name": "_xlnm.Print_Area", "localSheetId": first_index + offset})
        print_end = "AA" if remove_iwato else "AJ"
        area.appendChild(book.workbook.createTextNode(f"'{name}'!$A$1:${print_end}${len(selected_rows)}"))
        names.appendChild(area)
        book.files[path] = _daily_xml(source, selected_rows, shared)
        if remove_iwato:
            book.files[path] = _compact_layout(book, book.files[path], equalize_widths)
        if include_stamps:
            selected_stamps = stamp_groups[plan_dates[name]][:1] if remove_iwato else stamp_groups[plan_dates[name]]
            stamp_copies += _add_stamps(book, path, sheet_id, plan_dates[name], selected_stamps, right_align=remove_iwato, source=source)
    return book.finish(first_index), {"days": len(dates), "select_days": b_count, "stamps_removed": stamps if not include_stamps and not remove_iwato else 0, "stamps_kept": stamps if include_stamps or remove_iwato else 0, "stamp_copies": stamp_copies, "include_stamps": bool(include_stamps), "remove_iwato": bool(remove_iwato), "equalize_widths": bool(equalize_widths), "old_sheets_removed": removed, "sheets": [name for name, _ in plans]}


def cleanup_daily_sheets(data):
    """日別_で始まるシートだけを削除したダウンロード用ブックを返す。"""
    book = _Package(data)
    removed = book.remove_daily()
    return book.finish(), removed
