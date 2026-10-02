import zipfile
from io import BytesIO
from xml.sax.saxutils import escape
import json
import re
import unicodedata


def word_run(text, style=None):
    style = style or {}
    color = style.get("color", "#111827").lstrip("#")
    bold = "<w:b/>" if style.get("bold") else ""
    return f'<w:r><w:rPr>{bold}<w:color w:val="{escape(color)}"/></w:rPr><w:t xml:space="preserve">{escape(str(text))}</w:t></w:r>'


def word_paragraph(text, left=0, style=None, align=None):
    align_xml = f'<w:jc w:val="{align}"/>' if align else ""
    left_twips = max(0, int(float(left or 0) * 10))
    return f'<w:p><w:pPr><w:ind w:left="{left_twips}"/>{align_xml}</w:pPr>{word_run(text or " ", style)}</w:p>'


def word_table(table_grid, table_cells):
    if not table_grid or not table_grid.get("rows") or not table_grid.get("columns"):
        return ""

    cell_map = {(cell.get("row"), cell.get("col")): cell for cell in table_cells or []}
    rows_xml = []
    for row_index, row_height in enumerate(table_grid["rows"]):
        cells_xml = []
        for col_index, col_width in enumerate(table_grid["columns"]):
            cell = cell_map.get((row_index, col_index), {})
            lines = cell.get("lines") or ([cell.get("text", "")] if cell.get("text") else [""])
            paragraphs = "".join(
                word_paragraph(line, style=cell.get("style"), align="center")
                for line in lines
            )
            width_twips = max(120, int(float(col_width) * 12))
            cells_xml.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{width_twips}" w:type="dxa"/></w:tcPr>{paragraphs}</w:tc>'
            )
        rows_xml.append(f'<w:tr>{"".join(cells_xml)}</w:tr>')

    return (
        '<w:tbl><w:tblPr><w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="111827"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="111827"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="111827"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="111827"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="111827"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="111827"/>'
        f'</w:tblBorders></w:tblPr>{"".join(rows_xml)}</w:tbl>'
    )


def build_docx_paragraphs(text, formatted_data=None):
    formatted_data = formatted_data or {}
    layout_rows = formatted_data.get("layout_rows") or []
    table_grid = formatted_data.get("table_grid")
    table_cells = formatted_data.get("table_cells") or []

    if layout_rows or table_grid:
        entries = [
            {"type": "row", "y": float(row.get("y") or 0), "row": row}
            for row in layout_rows
        ]
        if table_grid:
            entries.append({"type": "table", "y": float(table_grid.get("y") or 0)})
        entries.sort(key=lambda entry: entry["y"])

        body_parts = []
        for entry in entries:
            if entry["type"] == "table":
                body_parts.append(word_table(table_grid, table_cells))
                continue

            row = entry["row"]
            align = "center" if (row.get("style") or {}).get("align") == "center" else None
            body_parts.append(
                word_paragraph(
                    row.get("text", ""),
                    left=row.get("x", 0),
                    style=row.get("style"),
                    align=align,
                )
            )
        return "".join(body_parts)

    return "".join(word_paragraph(line) for line in str(text or "").splitlines())


def make_docx(text, formatted_data=None, pages=None):
    """Build a Word file. If pages has multiple items, each becomes a Word page."""
    if pages and len(pages) > 1:
        body_parts = []
        for index, page in enumerate(pages):
            if index > 0:
                body_parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            title = page.get("title") or f"Trang {index + 1}"
            body_parts.append(word_paragraph(title, style={"bold": True}, align="center"))
            body_parts.append(
                build_docx_paragraphs(page.get("text", ""), page.get("formatted_data") or {})
            )
        paragraphs = "".join(body_parts)
    else:
        paragraphs = build_docx_paragraphs(text, formatted_data)

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {paragraphs}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document_xml)
    return output.getvalue()


def build_xlsx_sheet_rows(text, table_rows=None):
    rows = []
    source_rows = table_rows if table_rows else [[line] for line in str(text or "").splitlines()]
    for row_number, source_row in enumerate(source_rows, start=1):
        cells = []
        for column_number, cell in enumerate(source_row, start=1):
            column_name = xlsx_column_name(column_number)
            cells.append(
                f'<c r="{column_name}{row_number}" t="inlineStr"><is><t>{escape(str(cell))}</t></is></c>'
            )
        rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')
    return "".join(rows)


def sanitize_xlsx_sheet_name(name, fallback="Trang"):
    cleaned = re.sub(r'[\\/*?:\[\]]', " ", str(name or fallback)).strip() or fallback
    return cleaned[:31]


def make_xlsx(text, table_rows=None, pages=None):
    """Build an Excel file. If pages has multiple items, each becomes its own sheet."""
    if pages and len(pages) > 1:
        sheet_defs = []
        used_names = set()
        for index, page in enumerate(pages, start=1):
            base_name = sanitize_xlsx_sheet_name(
                page.get("sheet_name") or page.get("title") or f"Trang {index}",
                fallback=f"Trang {index}",
            )
            sheet_name = base_name
            suffix = 2
            while sheet_name.lower() in used_names:
                trim = 31 - len(f" ({suffix})")
                sheet_name = f"{base_name[:trim]} ({suffix})"
                suffix += 1
            used_names.add(sheet_name.lower())
            sheet_defs.append(
                {
                    "name": sheet_name,
                    "rows_xml": build_xlsx_sheet_rows(page.get("text", ""), page.get("table_rows")),
                }
            )
    else:
        sheet_defs = [
            {
                "name": "OCR Result",
                "rows_xml": build_xlsx_sheet_rows(text, table_rows),
            }
        ]

    sheet_entries = []
    sheet_rels = []
    content_overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    ]
    sheet_files = {}
    for index, sheet in enumerate(sheet_defs, start=1):
        sheet_path = f"worksheets/sheet{index}.xml"
        sheet_entries.append(
            f'<sheet name="{escape(sheet["name"])}" sheetId="{index}" r:id="rId{index}"/>'
        )
        sheet_rels.append(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{sheet_path}"/>'
        )
        content_overrides.append(
            f'<Override PartName="/xl/{sheet_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        sheet_files[f"xl/{sheet_path}"] = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>{sheet["rows_xml"]}</sheetData>
</worksheet>"""

    workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{''.join(sheet_entries)}</sheets>
</workbook>"""
    workbook_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {''.join(sheet_rels)}
</Relationships>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {''.join(content_overrides)}
</Types>"""

    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types)
        xlsx.writestr("_rels/.rels", root_rels)
        xlsx.writestr("xl/workbook.xml", workbook_xml)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        for path, xml in sheet_files.items():
            xlsx.writestr(path, xml)
    return output.getvalue()


def xlsx_column_name(column_number):
    name = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        name = chr(65 + remainder) + name
    return name


def normalize_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def strip_diacritics(value):
    text = unicodedata.normalize("NFD", normalize_text(value))
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return text.replace("Đ", "D").replace("đ", "d")


def ocr_text_lines(text):
    return [
        re.sub(r"^\d+\s*[:\.]\s*", "", normalize_text(line))
        for line in str(text or "").splitlines()
        if normalize_text(line)
    ]


def row_by_key(rows, key):
    key = key.lower()
    for row in rows or []:
        if len(row) > 2 and normalize_text(row[2]).lower() == key:
            return row
    return None


def first_matching_row(rows, predicate):
    for row in rows or []:
        if any(predicate(normalize_text(cell)) for cell in row):
            return row
    return None


def is_full_name_label(value):
    text = normalize_text(value)
    return bool(
        re.search(
            r"(?:H[ọoọóòõ]\s*v[àaá]\s*t[êeéèẽẹ]n|"
            r"M[ọo]\s*v[àaá]\s*t[êeéèẽẹ]n|"
            r"Full\s*nam|"
            r"Fuli\s*nam|"
            r"Ho\s*va\s*ten|"
            r"v[àaá]\s*t[êeéèẽẹ]n)",
            text,
            re.I,
        )
    )


def is_person_name_value(value):
    text = normalize_text(value)
    if not text or len(text) < 5:
        return False
    if re.search(r"\d", text):
        return False
    if re.search(
        r"VI[ỆE]T\s*NAM|CĂN\s*CƯỚC|CỘNG\s*HÒA|SOCIALIST|Citizen|Identity|Quốc\s*tịch|Nationality|"
        r"Ngày\s*sinh|Date\s*of|Giới|Sex|Quê|Place|Nơi|residence|expiry|Hạnh\s*phúc",
        text,
        re.I,
    ):
        return False
    return bool(re.fullmatch(r"[A-ZÀ-ỴĐ][A-ZÀ-ỴĐ\s]{3,}", text))


def is_gender_label(value):
    text = normalize_text(value)
    return bool(re.search(r"Gi[ớơòóõọ]?i?\s*t[iíìĩị]?nh?|\bS[eE][xXr]\b", text, re.I))


def extract_gender(text):
    normalized = normalize_text(text)
    match = re.search(
        r"(?:S[eE][xXr]|Gi[ớơòóõọ]?i?\s*t[iíìĩị]?nh?)\s*[:/|]?\s*(Nam|Nữ|Nu|Male|Female)",
        normalized,
        re.I,
    )
    if not match:
        # Value may sit alone on the same/next OCR fragment.
        match = re.search(r"\b(Nam|Nữ|Nu|Male|Female)\b", normalized, re.I)
        if not match or not is_gender_label(normalized):
            return ""
    value = match.group(1).lower()
    if value == "male":
        return "Nam"
    if value in ("female", "nu"):
        return "Nữ"
    if value == "nam":
        return "Nam"
    if value == "nữ":
        return "Nữ"
    return match.group(1)


def next_non_empty_single_value(rows, start_index):
    for row in (rows or [])[start_index + 1:]:
        cells = [normalize_text(cell) for cell in row if normalize_text(cell)]
        if len(cells) == 1:
            return cells[0]
    return ""


def ignored_address_text(value):
    return (
        re.search(r"Có\s*giá\s*trị|Date\s*of\s*expiry|expiry|hạn\s*dùng|valid", value, re.I)
        or re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", value)
        or re.search(r"Place\s*of\s*residence|Nơi\s*thường\s*trú", value, re.I)
    )


def address_like(value):
    return (
        "," in value
        or re.search(r"Tỉnh|Thành|Huyện|Quận|Xã|Phường|Thôn|Xóm|Ấp\.?", value, re.I)
    )


def next_address_value(rows, start_index):
    for row in (rows or [])[start_index + 1:]:
        cells = [normalize_text(cell) for cell in row if normalize_text(cell)]
        if not cells:
            continue
        row_text = " ".join(cells)
        if ignored_address_text(row_text):
            continue
        for cell in cells:
            if not ignored_address_text(cell) and address_like(cell):
                return cell
        if address_like(row_text):
            return row_text
    return ""


def find_ocr_line(lines, predicate):
    return next((line for line in lines if predicate(line)), "")


def next_address_line_after_residence(lines):
    residence_index = next(
        (index for index, line in enumerate(lines) if re.search(r"Place\s*of\s*residence|Nơi\s*thường\s*trú", line, re.I)),
        -1,
    )
    if residence_index < 0:
        return ""
    for line in lines[residence_index + 1:]:
        if ignored_address_text(line):
            continue
        if address_like(line):
            return line
    return ""


def build_front_citizen_json(text, table_rows=None):
    rows = table_rows or []
    lines = ocr_text_lines(text)
    output = {
        "CitizenId": "",
        "OwnerName": "",
        "OwnerBirthYear": "",
        "OwnerGender": "",
        "Address": "",
    }

    citizen_id_row = row_by_key(rows, "CitizenId") or first_matching_row(
        rows,
        lambda cell: re.search(r"(?:^|\b)(?:Số|No\.?)\b", normalize_text(cell), re.I)
        and not re.search(r"thường\s*trú|residence|origin|quê", normalize_text(cell), re.I),
    )
    if citizen_id_row:
        candidate = normalize_text(
            next((cell for cell in citizen_id_row if re.fullmatch(r"\d{9,12}", normalize_text(cell))), "")
            or (citizen_id_row[1] if len(citizen_id_row) > 1 else "")
        )
        match = re.search(r"\b\d{9,12}\b", candidate)
        output["CitizenId"] = match.group(0) if match else ""
    if not output["CitizenId"]:
        line = find_ocr_line(lines, lambda value: re.search(r"\b\d{9,12}\b", value))
        match = re.search(r"\b\d{9,12}\b", line)
        output["CitizenId"] = match.group(0) if match else ""

    owner_name_row = row_by_key(rows, "OwnerName")
    if owner_name_row:
        candidate = normalize_text(owner_name_row[0])
        output["OwnerName"] = candidate if is_person_name_value(candidate) else normalize_text(
            next((cell for cell in owner_name_row if is_person_name_value(normalize_text(cell))), candidate)
        )
    else:
        label_index = next(
            (index for index, row in enumerate(rows) if any(is_full_name_label(cell) for cell in row)),
            -1,
        )
        if label_index >= 0:
            label_row = rows[label_index]
            same_row_name = next(
                (normalize_text(cell) for cell in label_row if is_person_name_value(normalize_text(cell))),
                "",
            )
            output["OwnerName"] = same_row_name or next_non_empty_single_value(rows, label_index)
            if output["OwnerName"] and not is_person_name_value(output["OwnerName"]):
                output["OwnerName"] = next(
                    (
                        normalize_text(cell)
                        for row in rows[label_index + 1 : label_index + 4]
                        for cell in row
                        if is_person_name_value(normalize_text(cell))
                    ),
                    output["OwnerName"] if is_person_name_value(output["OwnerName"]) else "",
                )
    if not output["OwnerName"]:
        name_label_index = next((index for index, line in enumerate(lines) if is_full_name_label(line)), -1)
        if name_label_index >= 0:
            output["OwnerName"] = next(
                (line for line in lines[name_label_index + 1 : name_label_index + 5] if is_person_name_value(line)),
                "",
            )
    if not output["OwnerName"]:
        # Fallback: uppercase personal name near CitizenId / before birth date.
        id_index = next((index for index, line in enumerate(lines) if re.fullmatch(r"\d{9,12}", line)), -1)
        search_from = id_index + 1 if id_index >= 0 else 0
        output["OwnerName"] = next(
            (line for line in lines[search_from : search_from + 8] if is_person_name_value(line)),
            "",
        )

    birth_row = row_by_key(rows, "OwnerBirthYear") or first_matching_row(rows, lambda cell: re.search(r"Ngày\s*sinh|Date\s*of\s*birth|Ngay\s*sinh|Done\s*of\s*bird", cell, re.I))
    if birth_row:
        candidate = normalize_text(
            next((cell for cell in birth_row if re.search(r"\d{1,2}/\d{1,2}/\d{4}", normalize_text(cell))), "")
            or (birth_row[1] if len(birth_row) > 1 else "")
            or " ".join(str(cell) for cell in birth_row)
        )
        match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", candidate)
        output["OwnerBirthYear"] = match.group(0) if match else ""
    if not output["OwnerBirthYear"]:
        label_index = next(
            (index for index, line in enumerate(lines) if re.search(r"Ngày\s*sinh|Date\s*of\s*birth|Ngay\s*sinh", line, re.I)),
            -1,
        )
        search_lines = lines[label_index:label_index + 3] if label_index >= 0 else lines
        for line in search_lines:
            match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", line)
            if match:
                output["OwnerBirthYear"] = match.group(0)
                break

    gender_row = row_by_key(rows, "OwnerGender") or first_matching_row(rows, lambda cell: is_gender_label(cell))
    if gender_row:
        output["OwnerGender"] = extract_gender(" ".join(str(cell) for cell in gender_row))
    if not output["OwnerGender"]:
        output["OwnerGender"] = extract_gender(find_ocr_line(lines, is_gender_label))

    address_row = row_by_key(rows, "Address")
    if address_row:
        output["Address"] = normalize_text((address_row[1] if len(address_row) > 1 else "") or next((cell for cell in address_row if address_like(normalize_text(cell))), ""))
    if not output["Address"]:
        residence_index = next(
            (index for index, row in enumerate(rows) if any(re.search(r"Place\s*of\s*residence|Nơi\s*thường\s*trú", normalize_text(cell), re.I) for cell in row)),
            -1,
        )
        if residence_index >= 0:
            output["Address"] = next_address_value(rows, residence_index)
    if not output["Address"]:
        output["Address"] = next_address_line_after_residence(lines)

    return output


def is_issue_date_label(value):
    return bool(
        re.search(
            r"Ngày\s*,?\s*tháng\s*,?\s*năm|Date\s*,?\s*month\s*,?\s*year|"
            r"Ngay\s*,?\s*thang\s*,?\s*nam",
            normalize_text(value),
            re.I,
        )
    )


def is_expiry_date_label(value):
    return bool(
        re.search(
            r"Ngày\s*,?\s*tháng\s*,?\s*năm\s*hết\s*hạn|"
            r"Ngay\s*,?\s*thang\s*,?\s*nam\s*het\s*han|"
            r"Date\s*of\s*Expiry",
            normalize_text(value),
            re.I,
        )
    )


def is_legacy_back_issue_anchor(value):
    return bool(
        re.search(
            r"Ngày\s*,?\s*tháng\s*,?\s*năm|Date\s*,?\s*month\s*,?\s*year|"
            r"Ngay\s*,?\s*thang\s*,?\s*nam",
            normalize_text(value),
            re.I,
        )
    )


def extract_issue_date(value):
    match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", normalize_text(value))
    return match.group(0) if match else ""


def is_issue_place_noise(value):
    text = normalize_text(value)
    return bool(
        re.search(
            r"DIRECTOR|ADMINISTRATIVE|POLICE\s*DEPARTMENT|SOCIAL\s*ORDER|"
            r"Ngón\s*trỏ|index\s*finger|IDVNM|<<|Personal\s*identification|"
            r"Đặc\s*điểm|Nốt\s*ruồi|Nót\s*ruối|fingerprint",
            text,
            re.I,
        )
        or re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", text)
        or len(text) <= 2
    )


def is_issue_place_line(value):
    text = normalize_text(value)
    if not text or is_issue_place_noise(text):
        return False
    return bool(
        re.search(
            r"CỤC|CUC|CẢNH\s*SÁT|CANH\s*SAT|CÀNH\s*SÁT|QUẢN\s*LÝ|QUAN\s*LY|QUẬN\s*LÝ|"
            r"TRẬT\s*TỰ|TRAT\s*TU|XÃ\s*HỘI|XA\s*HOI|HÀNH\s*CHÍNH|HANH\s*CHINH|"
            r"BỘ\s*CÔNG\s*AN|BO\s*CONG\s*AN|MINISTRY\s*OF\s*PUBLIC\s*SECURITY",
            text,
            re.I,
        )
        or (text.isupper() and len(text) >= 12 and re.search(r"CỤC|CUC|CÔNG\s*AN|CONG\s*AN|MINISTRY|POLICE", text, re.I))
    )


def extract_issue_place_from_lines(lines, start_index):
    place_parts = []
    for line in lines[start_index + 1 :]:
        if is_issue_date_label(line) or re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", line):
            continue
        if is_issue_place_noise(line):
            if place_parts:
                break
            continue
        if is_issue_place_line(line) or (place_parts and re.search(r"[A-ZÀ-ỴĐ]", line) and len(line) >= 8):
            place_parts.append(line)
            if len(place_parts) >= 3:
                break
            continue
        if place_parts:
            break
    return normalize_issue_place_vietnamese(" ".join(place_parts))


def extract_issue_place_after_rows(rows, anchor_predicate, lookahead=2):
    anchor_index = next(
        (index for index, row in enumerate(rows) if any(anchor_predicate(cell) for cell in row)),
        -1,
    )
    if anchor_index < 0:
        return ""

    values = []
    seen = 0
    for row in rows[anchor_index + 1 :]:
        row_text = normalize_text(" ".join(str(cell) for cell in row if normalize_text(cell)))
        if not row_text:
            continue
        seen += 1
        if is_issue_place_line(row_text):
            values.append(normalize_issue_place_vietnamese(row_text))
        if seen >= lookahead:
            break
    return normalize_issue_place_vietnamese(" ".join(values))


def extract_issue_place_after_lines(lines, anchor_predicate, lookahead=2):
    anchor_index = next((index for index, line in enumerate(lines) if anchor_predicate(line)), -1)
    if anchor_index < 0:
        return ""

    values = []
    seen = 0
    for line in lines[anchor_index + 1 :]:
        if not normalize_text(line):
            continue
        seen += 1
        if is_issue_place_line(line):
            values.append(normalize_issue_place_vietnamese(line))
        if seen >= lookahead:
            break
    return normalize_issue_place_vietnamese(" ".join(values))


def normalize_issue_place_vietnamese(value):
    text = normalize_text(value)
    if not text:
        return ""

    segments = [segment.strip(" -:") for segment in re.split(r"\s*/\s*", text) if segment.strip(" -:")]
    if len(segments) > 1:
        preferred_segment = next(
            (
                segment
                for segment in segments
                if re.search(r"[À-Ỵà-ỵĐđ]", segment)
                or re.search(
                    r"\b(BỘ|CỤC|CẢNH|QUẢN|TRẬT|HÀNH|CÔNG|B[OỘ]|C[UỤ]C|CANH|QUAN|TRAT|HANH|CONG)\b",
                    segment,
                    re.I,
                )
            ),
            "",
        )
        if preferred_segment:
            text = preferred_segment

    ascii_key = re.sub(r"[^A-Z0-9 ]+", " ", strip_diacritics(text).upper())
    ascii_key = re.sub(r"\s+", " ", ascii_key).strip()

    if "BO CONG AN" in ascii_key:
        return "BỘ CÔNG AN"

    if (
        "CUC TRUONG CUC CANH SAT" in ascii_key
        and "QUAN LY HANH CHINH" in ascii_key
        and "TRAT TU XA HOI" in ascii_key
    ):
        return "CỤC TRƯỞNG CỤC CẢNH SÁT QUẢN LÝ HÀNH CHÍNH VỀ TRẬT TỰ XÃ HỘI"

    if (
        "CUC CANH SAT" in ascii_key
        and "QUAN LY HANH CHINH" in ascii_key
        and "TRAT TU XA HOI" in ascii_key
    ):
        return "CỤC CẢNH SÁT QUẢN LÝ HÀNH CHÍNH VỀ TRẬT TỰ XÃ HỘI"

    return text


def next_non_empty_values(rows, start_index, count):
    values = []
    for row in (rows or [])[start_index + 1:]:
        cells = [normalize_text(cell) for cell in row if normalize_text(cell)]
        if not cells:
            continue
        values.append(cells[0] if len(cells) == 1 else " ".join(cells))
        if len(values) >= count:
            break
    return values


def is_residence_label(value):
    text = normalize_text(value)
    return bool(
        re.search(
            r"N[ơo]i\s*c[ưu]\s*(?:tr[úuư]|ch[úu])|Place\s*of\s*residence",
            text,
            re.I,
        )
    )


def is_back_metadata_label(value):
    text = normalize_text(value)
    return bool(
        is_residence_label(text)
        or re.search(
            r"N[ơo]i\s*đ[ăa]ng\s*k[ýy]\s*khai\s*sinh|Place\s*of\s*birth|"
            r"Ng[àa]y\s*,?\s*th[áa]ng\s*,?\s*n[ăa]m\s*c[ấa]p|Date\s*of\s*issue|"
            r"Ng[àa]y\s*,?\s*th[áa]ng\s*,?\s*n[ăa]m\s*h[ếe]t\s*h[ạa]n|Date\s*of\s*Expiry|"
            r"B[ỘO]\s*C[ÔO]NG\s*AN|MINISTRY\s*OF\s*PUBLIC\s*SECURITY|"
            r"Ng[óo]n\s*tr[ỏo]|finger|Đ[ặa]c\s*đi[ểe]m|Personal\s*identification",
            text,
            re.I,
        )
    )


def extract_inline_value_after_colon(value):
    text = normalize_text(value)
    if ":" not in text:
        return ""
    return normalize_text(text.rsplit(":", 1)[-1])


def extract_residence_address_from_rows(rows):
    label_index = next(
        (index for index, row in enumerate(rows) if any(is_residence_label(cell) for cell in row)),
        -1,
    )
    if label_index < 0:
        return ""

    values = []
    saw_label = False
    for cell in rows[label_index]:
        cell_text = normalize_text(cell)
        if not cell_text:
            continue
        if is_residence_label(cell_text):
            saw_label = True
            inline_value = extract_inline_value_after_colon(cell_text)
            if inline_value and (address_like(inline_value) or not is_back_metadata_label(inline_value)):
                values.append(inline_value)
            continue
        if saw_label or address_like(cell_text) or values:
            values.append(cell_text)

    for row in rows[label_index + 1 :]:
        cells = [normalize_text(cell) for cell in row if normalize_text(cell)]
        if not cells:
            continue
        row_text = normalize_text(" ".join(cells))
        if is_back_metadata_label(row_text):
            break
        if ignored_address_text(row_text):
            if values:
                break
            continue
        if address_like(row_text) or values:
            values.append(row_text)
            if len(values) >= 3:
                break

    return normalize_text(" ".join(values))


def extract_residence_address_from_lines(lines):
    label_index = next((index for index, line in enumerate(lines) if is_residence_label(line)), -1)
    if label_index < 0:
        return ""

    values = []
    inline_value = extract_inline_value_after_colon(lines[label_index])
    if inline_value and (address_like(inline_value) or not is_back_metadata_label(inline_value)):
        values.append(inline_value)

    for line in lines[label_index + 1 :]:
        if is_back_metadata_label(line):
            break
        if ignored_address_text(line):
            if values:
                break
            continue
        if address_like(line) or values:
            values.append(line)
            if len(values) >= 3:
                break

    return normalize_text(" ".join(values))


def build_back_citizen_json(text, table_rows=None):
    rows = table_rows or []
    lines = ocr_text_lines(text)
    output = {
        "Address": "",
        "OwnerIssueDate": "",
        "OwnerIssuePlace": "",
    }

    issue_index = next(
        (index for index, row in enumerate(rows) if any(is_issue_date_label(cell) for cell in row)),
        -1,
    )
    if issue_index >= 0:
        issue_row = rows[issue_index]
        output["OwnerIssueDate"] = normalize_text(
            extract_issue_date(" ".join(str(cell) for cell in issue_row))
            or next(
                (
                    extract_issue_date(cell)
                    for cell in issue_row
                    if extract_issue_date(cell)
                ),
                "",
            )
        )
        place_start = issue_index
        if not output["OwnerIssueDate"]:
            for offset, row in enumerate(rows[issue_index + 1 :], start=1):
                cells = [normalize_text(cell) for cell in row if normalize_text(cell)]
                if not cells:
                    continue
                row_text = " ".join(cells)
                date_value = extract_issue_date(row_text)
                if date_value:
                    output["OwnerIssueDate"] = date_value
                    place_start = issue_index + offset
                break
        place_values = []
        for row in rows[place_start + 1 :]:
            row_text = normalize_text(" ".join(str(cell) for cell in row if normalize_text(cell)))
            if not row_text:
                continue
            if is_issue_place_noise(row_text):
                if place_values:
                    break
                continue
            if is_issue_place_line(row_text) or (place_values and len(row_text) >= 8):
                place_values.append(row_text)
                if len(place_values) >= 3:
                    break
                continue
            if place_values:
                break
        if place_values:
            output["OwnerIssuePlace"] = " ".join(place_values)

    if not output["OwnerIssuePlace"]:
        output["OwnerIssuePlace"] = (
            extract_issue_place_after_rows(rows, is_expiry_date_label, lookahead=2)
            or extract_issue_place_after_rows(rows, is_legacy_back_issue_anchor, lookahead=2)
        )

    label_index = next((index for index, line in enumerate(lines) if is_issue_date_label(line)), -1)
    if label_index >= 0:
        place_start = label_index
        if not output["OwnerIssueDate"]:
            output["OwnerIssueDate"] = extract_issue_date(lines[label_index])
            if not output["OwnerIssueDate"] and label_index + 1 < len(lines):
                next_date = extract_issue_date(lines[label_index + 1])
                if next_date:
                    output["OwnerIssueDate"] = next_date
                    place_start = label_index + 1
        if not output["OwnerIssuePlace"]:
            output["OwnerIssuePlace"] = extract_issue_place_from_lines(lines, place_start)

    if not output["OwnerIssuePlace"]:
        output["OwnerIssuePlace"] = (
            extract_issue_place_after_lines(lines, is_expiry_date_label, lookahead=2)
            or extract_issue_place_after_lines(lines, is_legacy_back_issue_anchor, lookahead=2)
        )

    if not output["OwnerIssueDate"]:
        for line in lines:
            date_value = extract_issue_date(line)
            if date_value and re.search(r"Ngày|Date|tháng|month|year", line, re.I):
                output["OwnerIssueDate"] = date_value
                break
        if not output["OwnerIssueDate"]:
            for line in lines:
                date_value = extract_issue_date(line)
                if date_value:
                    output["OwnerIssueDate"] = date_value
                    break

    if not output["OwnerIssuePlace"]:
        place_index = next((index for index, line in enumerate(lines) if is_issue_place_line(line)), -1)
        if place_index >= 0:
            output["OwnerIssuePlace"] = extract_issue_place_from_lines(lines, place_index - 1)

    output["Address"] = extract_residence_address_from_rows(rows)
    if not output["Address"]:
        output["Address"] = extract_residence_address_from_lines(lines)

    if output["OwnerIssuePlace"]:
        output["OwnerIssuePlace"] = normalize_issue_place_vietnamese(output["OwnerIssuePlace"])

    return output


def joined_ocr_text(text, table_rows=None):
    parts = ocr_text_lines(text)
    for row in table_rows or []:
        cells = [normalize_text(cell) for cell in row if normalize_text(cell)]
        if cells:
            parts.append(" ".join(cells))
    return "\n".join(parts)


def parse_vn_number(value):
    raw = normalize_text(value).replace(" ", "")
    if not raw:
        return ""
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?", raw):
        raw = raw.replace(".", "").replace(",", ".")
    elif "," in raw and "." not in raw:
        raw = raw.replace(",", ".")
    try:
        number = float(raw)
    except ValueError:
        return ""
    if number.is_integer():
        return str(int(number))
    return str(number)


def extract_row_value(rows, key):
    row = row_by_key(rows, key)
    if not row:
        return ""
    if len(row) > 1 and normalize_text(row[1]):
        return normalize_text(row[1])
    return normalize_text(row[0])


def extract_parcel_sheet_area(full_text):
    parcel = ""
    sheet = ""
    area = ""

    parcel_match = re.search(
        r"Th\S{0,3}a\s+đ\S{0,2}t\s+s\S{0,2}\s*[:\.]?\s*([0-9]+)",
        full_text,
        re.I,
    )
    if parcel_match:
        parcel = parcel_match.group(1)

    sheet_match = re.search(
        r"t\S{0,2}\s+b\S{0,2}n\s+đ\S{0,2}\s+s\S{0,2}\s*[:\.]?\s*([0-9]+)",
        full_text,
        re.I,
    )
    if sheet_match:
        sheet = sheet_match.group(1)

    area_match = re.search(
        r"Di\S{0,2}n\s+t\S{0,2}ch\s*[:\.]?\s*([0-9]+(?:[.,][0-9]+)*)\s*(?:m2|m²)?",
        full_text,
        re.I,
    )
    if area_match:
        area = parse_vn_number(area_match.group(1))

    return parcel, sheet, area


def extract_parcel_address(full_text, lines):
    patterns = (
        r"(?:^|\b)(?:[eе]\.|địa\s*ch[ỉi]|dia\s*chi)\s*[:\.]?\s*(.+)$",
        r"Địa\s*ch[ỉi]\s*th[ửứùa]\s*đ[ấa]t\s*[:\.]?\s*(.+)$",
    )
    for line in lines:
        for pattern in patterns:
            match = re.search(pattern, line, re.I)
            if not match:
                continue
            address = normalize_text(match.group(1))
            if address_like(address) and not re.search(r"tài\s*sản|tai\s*san", address, re.I):
                return address

    match = re.search(
        r"(?:[eе]\.|địa\s*ch[ỉi]|dia\s*chi)\s*[:\.]?\s*([^\n]+?(?:Tỉnh|Thành\s*phố|Huyện|Quận)[^\n]*)",
        full_text,
        re.I,
    )
    if match:
        return normalize_text(match.group(1))
    return ""


def extract_land_use_entries(full_text):
    section_match = re.search(
        r"Lo[ạa]i\s*đ[ấa]t\s*[:\.]?\s*(.+?)(?:\n|c\.\s*Th[ờơ]i\s*h[ạa]n|Th[ờơ]i\s*h[ạa]n\s*s[ửủ]\s*d[ụu]ng)",
        full_text,
        re.I | re.S,
    )
    section = normalize_text(section_match.group(1) if section_match else "")
    if not section:
        return [], []

    entries = re.findall(
        r"(Đ[ấa]t\s[^0-9:;,]{2,80}?)\s*([0-9]+(?:[.,][0-9]+)*)\s*(?:m2|m²)?",
        section,
        re.I,
    )
    land_uses = []
    areas = []
    for name, area in entries:
        cleaned = normalize_text(name).rstrip(" :;,-")
        if cleaned:
            land_uses.append(cleaned)
            areas.append(parse_vn_number(area))
    return land_uses, areas


def extract_owner_code(full_text, lines):
    for line in reversed(lines):
        match = re.search(r"\b([A-Z]{1,3}\s*\d{5,})\b", line)
        if match and not re.search(r"CCCD|Căn\s*cước", line, re.I):
            return normalize_text(match.group(1))

    match = re.search(r"\b([A-Z]{1,3}\s*\d{5,})\b", full_text)
    if match:
        return normalize_text(match.group(1))

    match = re.search(r"(?:s[ốô]\s*)?(?:CCCD|Căn\s*cước)\s*[:\.]?\s*(\d{9,12})", full_text, re.I)
    return match.group(1) if match else ""


def extract_land_type(full_text):
    match = re.search(
        r"(?:Đối\s*tượng\s*sử\s*dụng|Doi\s*tuong\s*su\s*dung)\s*[:\.]?\s*([^\n]+)",
        full_text,
        re.I,
    )
    if match:
        return normalize_text(match.group(1)).rstrip(" .;")

    if re.search(r"\b(Ông|Bà|Hộ\s*gia\s*đình)\b", full_text, re.I):
        return "Cá nhân"
    if re.search(r"\b(Công\s*ty|Tổ\s*chức|UBND|Ủy\s*ban)\b", full_text, re.I):
        return "Tổ chức"
    return ""


def extract_name_and_citizen_id(line):
    text = normalize_text(line)
    citizen_id = ""
    id_match = re.search(
        r"(?:s[ốô]\s*)?(?:CCCD|Căn\s*cước(?:\s*công\s*dân)?)\s*[:\.]?\s*(\d{9,12})",
        text,
        re.I,
    )
    if id_match:
        citizen_id = id_match.group(1)
    else:
        id_match = re.search(r"\b(\d{9,12})\b", text)
        if id_match:
            citizen_id = id_match.group(1)

    name = ""
    name_match = re.search(
        r"(?:Ông|Bà|Ông\s*bà|Hộ\s*ông|Hộ\s*bà)\s*[:\.]?\s*([^,;]+)",
        text,
        re.I,
    )
    if name_match:
        name = normalize_text(name_match.group(1))
        name = re.sub(
            r"\s*(?:s[ốô]\s*)?(?:CCCD|Căn\s*cước(?:\s*công\s*dân)?).*$",
            "",
            name,
            flags=re.I,
        ).strip(" :;-")
    return name, citizen_id


def extract_first_owner_identity(full_text, lines):
    candidates = list(lines)
    for line in full_text.splitlines():
        normalized = normalize_text(line)
        if normalized and normalized not in candidates:
            candidates.append(normalized)

    for line in candidates:
        if not re.search(r"\b(Ông|Bà)\b", line, re.I):
            continue
        name, citizen_id = extract_name_and_citizen_id(line)
        if name or citizen_id:
            return name, citizen_id
    return "", ""


def build_land_certificate_json(text, table_rows=None):
    rows = table_rows or []
    lines = ocr_text_lines(text)
    full_text = joined_ocr_text(text, rows)
    output = {
        "SheetNumber": "",
        "ParcelNumber": "",
        "Area": "",
        "ParcelAddress": "",
        "LandUse_GCN": "",
        "LandUse_Area": "",
        "OwnerCode": "",
        "LandType": "",
        "OwnerName": "",
        "CitizenId": "",
    }

    for key in output:
        value = extract_row_value(rows, key)
        if value:
            output[key] = value

    parcel, sheet, area = extract_parcel_sheet_area(full_text)
    if not output["ParcelNumber"]:
        output["ParcelNumber"] = parcel
    if not output["SheetNumber"]:
        output["SheetNumber"] = sheet
    if not output["Area"]:
        output["Area"] = area

    if not output["ParcelAddress"]:
        output["ParcelAddress"] = extract_parcel_address(full_text, lines)

    land_uses, land_areas = extract_land_use_entries(full_text)
    if not output["LandUse_GCN"] and land_uses:
        output["LandUse_GCN"] = "; ".join(land_uses)
    if not output["LandUse_Area"] and land_areas:
        output["LandUse_Area"] = "; ".join(land_areas)

    if not output["OwnerCode"]:
        output["OwnerCode"] = extract_owner_code(full_text, lines)
    if not output["LandType"]:
        output["LandType"] = extract_land_type(full_text)

    if not output["OwnerName"] or not output["CitizenId"]:
        owner_name, citizen_id = extract_first_owner_identity(full_text, lines)
        if not output["OwnerName"]:
            output["OwnerName"] = owner_name
        if not output["CitizenId"]:
            output["CitizenId"] = citizen_id

    return output


def empty_land_certificate_fields():
    return {
        "SheetNumber": "",
        "ParcelNumber": "",
        "Area": "",
        "ParcelAddress": "",
        "LandUse_GCN": "",
        "LandUse_Area": "",
        "OwnerCode": "",
        "LandType": "",
        "OwnerName": "",
        "CitizenId": "",
    }


def merge_land_certificate_fields(*parts):
    output = empty_land_certificate_fields()
    for part in parts:
        if not part:
            continue
        for key in output:
            value = normalize_text(part.get(key, ""))
            if value and not output[key]:
                output[key] = value
    return output


def build_land_certificate_template_2_json(text, table_rows=None):
    """Giấy CNSDĐ Mẫu số 2 (sổ hồng 2 trang) - same key_value schema as Mẫu số 1.

    Prefer merging extractors from Mẫu 1 + Mẫu 3 because page 1 usually has owners
    while page 2/3 has parcel/area/land-use details.
    """
    mau1 = build_land_certificate_json(text, table_rows)
    mau3 = build_land_certificate_template_3_json(text, table_rows)
    output = merge_land_certificate_fields(mau1, mau3)

    # Prefer first personal owner name from "Ông/Hộ ông" when Mau 1 picked a noisy value.
    lines = ocr_text_lines(text)
    mau3_name = extract_mau3_owner_name(lines)
    if mau3_name:
        output["OwnerName"] = mau3_name

    if not output["LandUse_GCN"]:
        output["LandUse_GCN"] = extract_mau3_land_use(joined_ocr_text(text, table_rows), lines)
    if not output["LandUse_Area"] and output["Area"]:
        output["LandUse_Area"] = output["Area"]
    if not output["LandType"]:
        output["LandType"] = extract_land_type(joined_ocr_text(text, table_rows)) or (
            "Cá nhân" if output["OwnerName"] else ""
        )
    if not output["OwnerCode"]:
        full_text = joined_ocr_text(text, table_rows)
        code_match = re.search(
            r"(?:S[ốô]\s*ph[áa]t\s*h[àa]nh\s*GCN|S[ốô]\s*v[àa]o\s*s[ổo]\s*c[ấa]p\s*GCN)\s*[:\.]?\s*([A-Z]{1,3}\s*\d{5,})",
            full_text,
            re.I,
        )
        if code_match:
            output["OwnerCode"] = normalize_text(code_match.group(1))
        else:
            output["OwnerCode"] = extract_owner_code(full_text, lines)
    return output


def next_value_after_label(lines, label_regex, value_regex=None, max_lookahead=3):
    for index, line in enumerate(lines):
        if not re.search(label_regex, line, re.I):
            continue
        inline = re.search(rf"{label_regex}\s*[:\.]?\s*(.+)$", line, re.I)
        if inline:
            candidate = normalize_text(inline.group(1))
            if candidate and (not value_regex or re.search(value_regex, candidate, re.I)):
                return candidate
        for offset in range(1, max_lookahead + 1):
            if index + offset >= len(lines):
                break
            candidate = normalize_text(lines[index + offset])
            if not candidate:
                continue
            if value_regex and not re.search(value_regex, candidate, re.I):
                continue
            return candidate
    return ""


def extract_mau3_parcel_sheet_area(full_text, lines):
    parcel = ""
    sheet = ""
    area = ""

    parcel_patterns = (
        r"Th\S{0,3}a\s+đ\S{0,2}t\s+s\S{0,2}\s*[:\.]?\s*([0-9]+[a-zA-Z]?)",
        r"a\s*\)?\s*Th\S{0,8}\s*[:\.]?\s*([0-9]+[a-zA-Z]?)",
        r"aThadsb\s*[:\.]?\s*([0-9]+[a-zA-Z]?)",
        r"Th\S{0,3}a\s*đ\S{0,2}t\s*s[ốôo]?\s*[:\.]?\s*([0-9]+[a-zA-Z]?)",
    )
    for pattern in parcel_patterns:
        parcel_match = re.search(pattern, full_text, re.I)
        if parcel_match:
            parcel = parcel_match.group(1)
            break
    if not parcel:
        parcel = next_value_after_label(
            lines,
            r"1\s*[,.]?\s*Th\S{0,3}a\s*đ\S{0,2}t|Th\S{0,3}a\s*đ\S{0,2}t\s*:",
            value_regex=r"^\d{1,4}[a-zA-Z]?$",
            max_lookahead=2,
        )

    sheet = next_value_after_label(
        lines,
        r"t\S{0,2}\s+b\S{0,2}n\s+đ\S{0,2}\s+s\S{0,2}",
        value_regex=r"^\d{1,4}$",
        max_lookahead=2,
    )
    if not sheet:
        sheet_match = re.search(
            r"(?:t\S{0,2}\s*)?b\S{0,2}n\s*đ\S{0,2}\s*s\S{0,2}\s*[:\.]?\s*([0-9]+)",
            full_text,
            re.I,
        )
        if sheet_match:
            sheet = sheet_match.group(1)
    # Same-line form: "Thửa đất số: 37 tờ bản đồ số: 4"
    if parcel and not sheet:
        same_line = re.search(
            rf"{re.escape(parcel)}\D{{0,40}}?(?:t\S{{0,2}}\s*)?b\S{{0,2}}n\s*đ\S{{0,2}}\s*s\S{{0,2}}\s*[:\.]?\s*([0-9]+)",
            full_text,
            re.I,
        )
        if same_line:
            sheet = same_line.group(1)

    area_match = re.search(
        r"Di\S{0,2}n\s*t\S{0,2}ch\s*[:\.]?\s*([0-9]+(?:[.,][0-9]+)*)",
        full_text,
        re.I,
    )
    if not area_match:
        area_match = re.search(
            r"(?:Điện\s*úch|Dintich\w*)\s*[:\.]?\s*([0-9]+(?:[.,][0-9]+)*)",
            full_text,
            re.I,
        )
    if not area_match:
        # Number often appears before the garbled "Diện tích" label, with "bằng chữ".
        area_match = re.search(
            r"\b([0-9]{2,4}(?:[.,][0-9]+)?)\s*m(?:2|²)?\.?\s*\([^)]*(?:chữ|chm|tram|trăm|mét|met)",
            full_text,
            re.I,
        )
    if area_match:
        area = parse_vn_number(area_match.group(1))
    if not area:
        private_match = re.search(
            r"S[ửủ]\s*d[ụu]ng\s*ri[êe]ng\s*[:\.]?\s*([0-9]+(?:[.,][0-9]+)*)",
            full_text,
            re.I,
        )
        if private_match:
            area = parse_vn_number(private_match.group(1))

    return parcel, sheet, area


def extract_mau3_parcel_address(full_text, lines):
    patterns = (
        r"Địa\s*ch[ỉi]\s*th\S{0,3}a\s+đ\S{0,2}t\s*[:\.]?\s*(.+)$",
        r"(?:b\s*\)?\s*)?(?:Địa\s*ch[ỉi]|Da\s*chi|Bị\s*Da\s*ch[ỉi])\s*[,:.]?\s*(.+)$",
        r"bDa\s*chi\s*(.+)$",
    )
    for line in lines:
        for pattern in patterns:
            match = re.search(pattern, line, re.I)
            if not match:
                continue
            address = normalize_text(match.group(1)).lstrip(" ,:.-")
            if address_like(address) or (
                len(address) >= 12
                and re.search(r"Phường|Quận|Huyện|Xã|Thành\s*phố|Tỉnh", address, re.I)
            ):
                return address

    match = re.search(
        r"(?:Địa\s*ch[ỉi]|Da\s*chi)\s*(?:th\S{0,3}a\s+đ\S{0,2}t)?\s*[:.,]?\s*"
        r"([^\n]+?(?:Tỉnh|Thành\s*phố|Hà\s*N[ộo]i)[^\n]*)",
        full_text,
        re.I,
    )
    if match:
        return normalize_text(match.group(1)).lstrip(" ,:.-")
    return ""


def extract_mau3_land_use(full_text, lines):
    land_use = next_value_after_label(
        lines,
        r"M[ụu]c\s*đ[íiì]ch\s*s[ửủ]\s*d[ụu]ng(?:\s*đ[ấaá]t)?",
        value_regex=r"Đ[ấaá]t\s+",
        max_lookahead=2,
    )
    if not land_use:
        match = re.search(
            r"M[ụu]c\s*d[íiì]ch\s*s[ửủ]\s*d[ụu]ng(?:\s*đ[ấaá]t)?\s*[:\.]?\s*([^\n]+)",
            full_text,
            re.I,
        )
        if match:
            land_use = normalize_text(match.group(1))
    if not land_use:
        match = re.search(
            r"(Đ[ấaá]t\s+[^\n,]{0,40}?(?:đô\s*thị|nông\s*thôn|ở\s+[^\n,]{2,30}))",
            full_text,
            re.I,
        )
        if match:
            land_use = normalize_text(match.group(1))
    land_use = re.sub(r"\s*Th[ờơ]i\s*h[ạa]n.*$", "", land_use or "", flags=re.I).strip(" :;-")
    # Common OCR drift on "Đất ở tại đô thị"
    if re.search(r"^D[ấaá]t\s+[oở].{0,20}th[ịi]\s*$", land_use or "", re.I):
        land_use = "Đất ở tại đô thị"
    return land_use


def extract_mau3_owner_code(full_text, lines):
    code = extract_owner_code(full_text, lines)
    if code:
        return code
    for line in reversed(lines):
        match = re.search(
            r"(?:S[ốô]\s*v[àa]o\s*s[ổô].{0,20}?)?(CS[- ]?B[T1I]L)\s*[:.]?\s*([0-9OIlSS]{4,})",
            line,
            re.I,
        )
        if match:
            digits = (
                match.group(2)
                .upper()
                .replace("O", "0")
                .replace("I", "1")
                .replace("L", "1")
                .replace("S", "5")
            )
            prefix = re.sub(r"[^A-Z]", "", match.group(1).upper()).replace("B1L", "BTL").replace("BIL", "BTL")
            if not prefix.startswith("CS"):
                prefix = "CS" + prefix
            if "BTL" not in prefix:
                prefix = "CS-BTL"
            else:
                prefix = "CS-BTL"
            return f"{prefix}: {digits}"
    match = re.search(
        r"(CS[- ]?B[T1I]L)\s*[:.]?\s*([0-9OIlSS]{4,})",
        full_text,
        re.I,
    )
    if match:
        digits = (
            match.group(2)
            .upper()
            .replace("O", "0")
            .replace("I", "1")
            .replace("L", "1")
            .replace("S", "5")
        )
        return f"CS-BTL: {digits}"
    return ""


def extract_mau3_owner_name(lines):
    owner_label = r"(?:H[ộo]\s*[oôóòọ]ng|^\s*Ông|^\s*Bà)\b"
    for index, line in enumerate(lines):
        if not re.search(owner_label, line, re.I):
            continue
        inline = re.search(r"(?:H[ộo]\s*[oôóòọ]ng|Ông|Bà)\s*[:;,]?\s*(.+)$", line, re.I)
        if inline and normalize_text(inline.group(1)):
            name = normalize_text(inline.group(1))
            name = re.sub(r"\s*Năm\s*sinh.*$", "", name, flags=re.I).strip(" :;-")
            if name and not re.search(r"Năm\s*sinh|Địa\s*ch[ỉi]", name, re.I):
                return name
        for offset in range(1, 3):
            if index + offset >= len(lines):
                break
            candidate = normalize_text(lines[index + offset])
            if not candidate:
                continue
            if re.search(r"Năm\s*sinh|Địa\s*ch[ỉi]|^\s*Bà\b|^\s*Ông\b|H[ộo]\s*[oôóòọ]ng", candidate, re.I):
                continue
            if re.search(r"\d{4}", candidate):
                continue
            if len(candidate) >= 5:
                return candidate
    return ""


def build_land_certificate_template_3_json(text, table_rows=None):
    """Giấy CNSDĐ Mẫu số 3 - same key_value schema as Mẫu số 1."""
    rows = table_rows or []
    lines = ocr_text_lines(text)
    full_text = joined_ocr_text(text, rows)
    output = {
        "SheetNumber": "",
        "ParcelNumber": "",
        "Area": "",
        "ParcelAddress": "",
        "LandUse_GCN": "",
        "LandUse_Area": "",
        "OwnerCode": "",
        "LandType": "",
        "OwnerName": "",
        "CitizenId": "",
    }

    for key in output:
        value = extract_row_value(rows, key)
        if value:
            output[key] = value

    parcel, sheet, area = extract_mau3_parcel_sheet_area(full_text, lines)
    if not output["ParcelNumber"]:
        output["ParcelNumber"] = parcel
    if not output["SheetNumber"]:
        output["SheetNumber"] = sheet
    if not output["Area"]:
        output["Area"] = area

    if not output["ParcelAddress"]:
        output["ParcelAddress"] = extract_mau3_parcel_address(full_text, lines)
        # Fallback to Mẫu 1 style address labels on pink-book page 2.
        if not output["ParcelAddress"]:
            output["ParcelAddress"] = extract_parcel_address(full_text, lines)

    if not output["LandUse_GCN"]:
        output["LandUse_GCN"] = extract_mau3_land_use(full_text, lines)
    if not output["LandUse_Area"]:
        output["LandUse_Area"] = output["Area"]

    if not output["OwnerCode"]:
        output["OwnerCode"] = extract_mau3_owner_code(full_text, lines)
    if not output["LandType"]:
        if re.search(r"H[ộo]\s*[oôóòọ]ng|\bÔng\b|\bBà\b", full_text, re.I):
            output["LandType"] = "Cá nhân"
        elif re.search(r"Căn\s*h[ộo]\s*chung\s*cư|chung\s*cư", full_text, re.I):
            output["LandType"] = "Cá nhân"
        else:
            output["LandType"] = extract_land_type(full_text)

    if not output["OwnerName"]:
        output["OwnerName"] = extract_mau3_owner_name(lines)
    if not output["CitizenId"]:
        _, citizen_id = extract_first_owner_identity(full_text, lines)
        output["CitizenId"] = citizen_id

    return output


def build_json(text, table_rows=None, json_type="cccd_front"):
    if json_type == "cccd_back":
        return build_back_citizen_json(text, table_rows)
    if json_type in ("land_certificate", "land_certificate_1"):
        return build_land_certificate_json(text, table_rows)
    if json_type == "land_certificate_2":
        return build_land_certificate_template_2_json(text, table_rows)
    if json_type == "land_certificate_3":
        return build_land_certificate_template_3_json(text, table_rows)
    return build_front_citizen_json(text, table_rows)


def make_json(text, table_rows=None, json_type="cccd_front"):
    return json.dumps(build_json(text, table_rows, json_type), ensure_ascii=False, indent=2).encode("utf-8")
