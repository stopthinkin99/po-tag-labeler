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


BOLD_FONT = "Helvetica-Bold"
NORMAL_FONT = "Helvetica"


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


def _find_safe_top_box(page: fitz.Page, box_w: float, box_h: float) -> tuple[float, float]:
    """Find a small blank spot near the top without covering existing PDF text."""
    page_rect = page.rect

    occupied = []
    for word in page.get_text("words"):
        x0, y0, x1, y1 = word[:4]
        occupied.append(fitz.Rect(x0 - 1, y0 - 1, x1 + 1, y1 + 1))

    x_centers = [0.50, 0.60, 0.40, 0.70, 0.30, 0.80, 0.20]
    y_values = [4, 12, 20, 28, 36, 44, 52, 60, 68, 76, 84]
    best: tuple[float, float, float] | None = None

    for y0 in y_values:
        for center in x_centers:
            x0 = page_rect.width * center - box_w / 2
            x0 = min(max(4, x0), page_rect.width - box_w - 4)
            candidate = fitz.Rect(x0, y0, x0 + box_w, y0 + box_h)
            overlap = sum(_overlap_area(candidate, rect) for rect in occupied)
            score = overlap * 1000 + y0 + abs(center - 0.5) * 10
            if best is None or score < best[0]:
                best = (score, x0, y0)
            if overlap == 0:
                return x0, y0

    assert best is not None
    return best[1], best[2]


def _remove_old_tag_box(page: fitz.Page) -> None:
    """Remove an older large 'TAG NO' stamp if the input PDF already has one."""
    matches = []
    for needle in ("TAG NO:", "TAG NO", "TAG"):
        try:
            matches.extend(page.search_for(needle))
        except Exception:
            pass

    candidates = [r for r in matches if r.y0 < 120]
    if not candidates:
        return

    rect = candidates[0]
    for r in candidates[1:]:
        rect |= r

    cleanup = fitz.Rect(
        max(page.rect.x0, rect.x0 - 85),
        max(page.rect.y0, rect.y0 - 18),
        min(page.rect.x1, rect.x1 + 85),
        min(page.rect.y1, rect.y1 + 20),
    )
    page.draw_rect(
        cleanup,
        color=(1, 1, 1),
        fill=(1, 1, 1),
        width=0,
        overlay=True,
    )


def _stamp_tag(
    page: fitz.Page,
    tags: Iterable[str],
    position: str = "auto-safe-top",
    font_size: int = 10,
) -> None:
    tag_text = ", ".join(tags)
    page_rect = page.rect
    actual_size = max(7, min(font_size, 16))
    text_width = fitz.get_text_length(tag_text, fontname=BOLD_FONT, fontsize=actual_size)

    circle_padding = max(5, actual_size * 0.55)
    diameter = max(text_width + circle_padding * 2, actual_size * 2.4)
    box_w = diameter
    box_h = diameter
    margin = 6

    if position == "auto-safe-top":
        x0, y0 = _find_safe_top_box(page, box_w, box_h)
    else:
        if position == "top-left":
            x0 = margin
        elif position == "top-center":
            x0 = (page_rect.width - box_w) / 2
        else:
            x0 = page_rect.width - box_w - margin
        y0 = margin

    center_x = x0 + diameter / 2
    center_y = y0 + diameter / 2
    radius = diameter / 2

    page.draw_circle(
        (center_x, center_y),
        radius,
        color=(0, 0, 0),
        width=1.3,
        fill=None,
        overlay=True,
    )

    baseline_y = center_y + (actual_size * 0.36)
    text_x = center_x - text_width / 2

    page.insert_text(
        (text_x, baseline_y),
        tag_text,
        fontsize=actual_size,
        fontname=BOLD_FONT,
        color=(0, 0, 0),
        overlay=True,
    )


def process_pdf(
    pdf_bytes: bytes,
    mapping: dict[str, list[str]],
    position: str = "auto-safe-top",
    font_size: int = 10,
    remove_old_tag_box: bool = True,
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

        if remove_old_tag_box:
            _remove_old_tag_box(page)

        _stamp_tag(page, tags, position=position, font_size=font_size)
        results.append(PageResult(index + 1, po, tags, "Matched"))

    output = document.tobytes(garbage=4, deflate=True)
    document.close()
    return output, results
