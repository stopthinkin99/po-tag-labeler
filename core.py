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


def _overlap_area(a: fitz.Rect, b: fitz.Rect) -> float:
    inter = a & b
    if inter.is_empty:
        return 0.0
    return max(0.0, inter.width) * max(0.0, inter.height)


def _find_safe_top_position(page: fitz.Page, text: str, font_size: float) -> tuple[float, float]:
    """Find a small blank spot near the top without covering existing PDF text."""
    page_rect = page.rect
    text_width = fitz.get_text_length(text, fontname="helv", fontsize=font_size)
    pad_x, pad_y = 3, 2
    box_w = text_width + pad_x * 2
    box_h = font_size + pad_y * 2

    occupied = []
    for word in page.get_text("words"):
        x0, y0, x1, y1 = word[:4]
        occupied.append(fitz.Rect(x0 - 1, y0 - 1, x1 + 1, y1 + 1))

    # Prefer the upper middle, then move across/down until a clear spot is found.
    x_centers = [0.50, 0.60, 0.40, 0.70, 0.30, 0.80, 0.20]
    y_values = [4, 12, 20, 28, 36, 44, 52, 60, 68, 76, 84]
    best: tuple[float, float, float] | None = None

    for y0 in y_values:
        for center in x_centers:
            x0 = page_rect.width * center - box_w / 2
            x0 = min(max(4, x0), page_rect.width - box_w - 4)
            candidate = fitz.Rect(x0, y0, x0 + box_w, y0 + box_h)
            overlap = sum(_overlap_area(candidate, rect) for rect in occupied)
            # Slight preference for higher positions and the horizontal center.
            score = overlap * 1000 + y0 + abs(center - 0.5) * 10
            if best is None or score < best[0]:
                best = (score, x0 + pad_x, y0 + pad_y + font_size)
            if overlap == 0:
                return x0 + pad_x, y0 + pad_y + font_size

    assert best is not None
    return best[1], best[2]


def _stamp_tag(
    page: fitz.Page,
    tags: Iterable[str],
    position: str = "auto-safe-top",
    font_size: int = 10,
) -> None:
    # Print only the tag value(s), not "TAG NO:" and not a large border box.
    tag_text = ", ".join(tags)
    page_rect = page.rect
    actual_size = max(7, min(font_size, 16))
    text_width = fitz.get_text_length(tag_text, fontname="helv", fontsize=actual_size)
    margin = 6

    if position == "auto-safe-top":
        text_x, text_y = _find_safe_top_position(page, tag_text, actual_size)
    else:
        if position == "top-left":
            text_x = margin
        elif position == "top-center":
            text_x = (page_rect.width - text_width) / 2
        else:
            text_x = page_rect.width - text_width - margin
        text_y = margin + actual_size

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
    position: str = "auto-safe-top",
    font_size: int = 10,
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
