from __future__ import annotations

import io
import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Iterable

import fitz  # PyMuPDF
from openpyxl import load_workbook


PO_HEADER_CANDIDATES = {
    "cust po no", "cust po number", "customer po", "customer po no",
    "customer po number", "po no", "po number",
}
TAG_HEADER_CANDIDATES = {
    "tag no", "tag number", "tag", "tag #",
}


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").strip().lower().split())


def normalize_po(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    text = re.sub(r"\.0$", "", text)
    return re.sub(r"[^A-Za-z0-9]", "", text).upper()


def load_po_tag_map(excel_bytes: bytes) -> dict[str, list[str]]:
    workbook = load_workbook(io.BytesIO(excel_bytes), data_only=True, read_only=True)

    for sheet in workbook.worksheets:
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            continue

        # Search the first 20 rows for a header row.
        for row_index, row in enumerate(rows[:20]):
            normalized = [normalize_header(cell) for cell in row]
            po_col = next((i for i, h in enumerate(normalized) if h in PO_HEADER_CANDIDATES), None)
            tag_col = next((i for i, h in enumerate(normalized) if h in TAG_HEADER_CANDIDATES), None)

            if po_col is None or tag_col is None:
                continue

            mapping: dict[str, list[str]] = defaultdict(list)
            for data_row in rows[row_index + 1:]:
                if po_col >= len(data_row) or tag_col >= len(data_row):
                    continue
                po = normalize_po(data_row[po_col])
                tag_value = data_row[tag_col]
                if not po or tag_value is None or str(tag_value).strip() == "":
                    continue
                if isinstance(tag_value, float) and tag_value.is_integer():
                    tag_value = int(tag_value)
                tag = str(tag_value).strip()
                if tag not in mapping[po]:
                    mapping[po].append(tag)

            if mapping:
                return dict(mapping)

    raise ValueError(
        "Could not find the Excel columns. Expected headers similar to "
        "'Cust Po No' and 'Tag NO'."
    )


def extract_po_from_page_text(text: str) -> str | None:
    # Current FedEx labels store the customer PO in REF. Some other labels may
    # have it in PO, so both are supported.
    patterns = [
        r"\bREF\s*:\s*([A-Z0-9][A-Z0-9\-/]*)",
        r"\bPO\s*(?:NO|NUMBER|#)?\s*:\s*([A-Z0-9][A-Z0-9\-/]*)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return normalize_po(match.group(1))
    return None


@dataclass
class PageResult:
    page_number: int
    po_number: str | None
    tags: list[str]
    status: str


def _stamp_tag(
    page: fitz.Page,
    tags: Iterable[str],
    position: str = "top-right",
    font_size: int = 18,
) -> None:
    tag_text = "TAG NO: " + ", ".join(tags)
    page_rect = page.rect
    box_width = min(max(170, 11 * len(tag_text)), page_rect.width - 24)
    box_height = max(34, font_size + 16)
    margin = 12

    if position == "top-left":
        x0 = margin
    elif position == "top-center":
        x0 = (page_rect.width - box_width) / 2
    else:
        x0 = page_rect.width - box_width - margin

    y0 = margin
    rect = fitz.Rect(x0, y0, x0 + box_width, y0 + box_height)

    # White background keeps the tag readable without touching barcodes lower
    # on the shipping label.
    page.draw_rect(rect, color=(0, 0, 0), fill=(1, 1, 1), width=1.2, overlay=True)
    text_width = fitz.get_text_length(tag_text, fontname="helv", fontsize=font_size)
    usable_width = rect.width - 12
    actual_size = font_size
    if text_width > usable_width:
        actual_size = max(8, font_size * usable_width / text_width)
        text_width = fitz.get_text_length(tag_text, fontname="helv", fontsize=actual_size)

    text_x = rect.x0 + max(6, (rect.width - text_width) / 2)
    text_y = rect.y0 + (rect.height + actual_size) / 2 - 3
    page.insert_text(
        (text_x, text_y),
        tag_text,
        fontsize=actual_size,
        fontname="helv",
        color=(0, 0, 0),
        overlay=True,
    )


def process_pdf(
    pdf_bytes: bytes,
    mapping: dict[str, list[str]],
    position: str = "top-right",
    font_size: int = 18,
) -> tuple[bytes, list[PageResult]]:
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: list[PageResult] = []

    for index, page in enumerate(document):
        page_text = page.get_text("text") or ""
        po = extract_po_from_page_text(page_text)

        if not po:
            results.append(PageResult(index + 1, None, [], "PO/REF not found"))
            continue

        tags = mapping.get(po, [])
        if not tags:
            results.append(PageResult(index + 1, po, [], "No Excel match"))
            continue

        _stamp_tag(page, tags, position=position, font_size=font_size)
        results.append(PageResult(index + 1, po, tags, "Matched"))

    output = document.tobytes(garbage=4, deflate=True)
    document.close()
    return output, results
