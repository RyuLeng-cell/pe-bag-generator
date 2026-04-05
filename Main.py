import os
import re
import sys
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


INPUT_FILE = "input.xlsx"
SHEET_NAME = "Sheet1"
OUTPUT_DIR = "output"
LOGO_FILE = "assets/logo.png"


def get_base_dir() -> Path:
    return Path(os.path.dirname(os.path.abspath(sys.argv[0])))


def show_info(title: str, msg: str):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(title, msg)
    root.destroy()


def show_error(title: str, msg: str):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showerror(title, msg)
    root.destroy()


def safe_filename(text: str, fallback: str = "output") -> str:
    if text is None:
        return fallback
    text = str(text).strip()
    if not text:
        return fallback
    text = re.sub(r'[\\/*?:"<>|]', "_", text)
    text = re.sub(r"\s+", "_", text)
    return text


def to_float(value, field_name: str) -> float:
    try:
        return float(value)
    except Exception:
        raise ValueError(f"Invalid numeric value for '{field_name}': {value}")


def clean_str(value, default: str = "") -> str:
    if pd.isna(value):
        return default
    return str(value).strip()


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    original_cols = list(df.columns)

    rename_map = {}
    for c in original_cols:
        key = str(c).strip().lower()
        rename_map[c] = key

    df = df.rename(columns=rename_map)

    required = {
        "material",
        "thickness",
        "length",
        "width",
        "hex",
        "seal width",
        "distance from seal",
        "program",
    }

    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(sorted(missing))
            + "\nFound columns: "
            + ", ".join(map(str, original_cols))
        )

    # Optional columns for public version
    optional_defaults = {
        "anti-static": "Required",
        "cleanroom grade": "Yes",
        "color": "",
        "date": "",
        "dwg no": "",
        "doc no": "",
    }

    for col, default_val in optional_defaults.items():
        if col not in df.columns:
            df[col] = default_val

    return df


def draw_text_center(c, x_center, y, text, font_name="Helvetica", font_size=10):
    c.setFont(font_name, font_size)
    tw = stringWidth(str(text), font_name, font_size)
    c.drawString(x_center - tw / 2, y, str(text))


def draw_box(c, x, y, w, h, stroke=1, fill=0):
    c.rect(x, y, w, h, stroke=stroke, fill=fill)


def draw_cell_text(
    c,
    x,
    y,
    w,
    h,
    text,
    align="left",
    font_name="Helvetica",
    font_size=10,
    padding=2 * mm,
):
    c.setFont(font_name, font_size)
    text = "" if text is None else str(text)

    baseline_adjust = font_size * 0.28
    text_y = y + (h / 2.0) - baseline_adjust

    if align == "center":
        tw = stringWidth(text, font_name, font_size)
        text_x = x + (w - tw) / 2.0
    elif align == "right":
        tw = stringWidth(text, font_name, font_size)
        text_x = x + w - padding - tw
    else:
        text_x = x + padding

    c.drawString(text_x, text_y, text)


def draw_dimension_line(c, x1, y1, x2, y2, text, text_offset=4 * mm, vertical=False):
    dim_color = colors.HexColor("#165A8A")
    c.setStrokeColor(dim_color)
    c.setFillColor(dim_color)
    c.setLineWidth(1)

    arrow = 2.2 * mm

    if not vertical:
        c.line(x1, y1, x2, y2)
        c.line(x1, y1, x1 + arrow, y1 + arrow / 2)
        c.line(x1, y1, x1 + arrow, y1 - arrow / 2)
        c.line(x2, y2, x2 - arrow, y2 + arrow / 2)
        c.line(x2, y2, x2 - arrow, y2 - arrow / 2)

        c.setFont("Helvetica", 11)
        tw = stringWidth(text, "Helvetica", 11)
        tx = (x1 + x2) / 2 - tw / 2
        ty = y1 - text_offset
        c.drawString(tx, ty, text)
    else:
        c.line(x1, y1, x2, y2)
        c.line(x1, y1, x1 - arrow / 2, y1 + arrow)
        c.line(x1, y1, x1 + arrow / 2, y1 + arrow)
        c.line(x2, y2, x2 - arrow / 2, y2 - arrow)
        c.line(x2, y2, x2 + arrow / 2, y2 - arrow)

        c.saveState()
        c.setFont("Helvetica", 11)
        tx = x1 + text_offset
        ty = (y1 + y2) / 2
        c.translate(tx, ty)
        c.rotate(90)
        c.drawString(0, 0, text)
        c.restoreState()


def draw_revision_table(c, x, y_top, w, row_h=8 * mm, data_rows=4):
    total_rows = 1 + data_rows
    total_h = total_rows * row_h
    y_bottom = y_top - total_h

    col_widths = [
        12 * mm,
        35 * mm,
        92 * mm,
        22 * mm,
        w - (12 * mm + 35 * mm + 92 * mm + 22 * mm),
    ]

    draw_box(c, x, y_bottom, w, total_h)

    current_x = x
    for cw in col_widths[:-1]:
        current_x += cw
        c.line(current_x, y_bottom, current_x, y_top)

    for i in range(1, total_rows):
        yy = y_top - i * row_h
        c.line(x, yy, x + w, yy)

    headers = ["REV", "DATE", "DESCRIPTION", "ECO NO", "SIGN"]
    cx = x
    header_y = y_top - row_h
    for header, cw in zip(headers, col_widths):
        draw_cell_text(
            c,
            cx,
            header_y,
            cw,
            row_h,
            header,
            align="center",
            font_name="Helvetica-Bold",
            font_size=10,
        )
        cx += cw


def draw_approval_block(c, x, y, w, h):
    draw_box(c, x, y, w, h)
    row_h = h / 3.0
    split_w = 28 * mm
    split_x = x + split_w

    c.line(split_x, y, split_x, y + h)
    c.line(x, y + row_h, x + w, y + row_h)
    c.line(x, y + 2 * row_h, x + w, y + 2 * row_h)

    labels = ["DESIGNED", "CHECKED", "APPROVAL"]
    for i, label in enumerate(labels):
        cell_y = y + h - (i + 1) * row_h
        draw_cell_text(c, x, cell_y, split_w, row_h, label, align="left", font_size=10)


def draw_logo_only_block(c, x, y, w, h, logo_path):
    draw_box(c, x, y, w, h)

    if os.path.exists(logo_path):
        try:
            img = ImageReader(logo_path)
            iw, ih = img.getSize()

            max_w = w - 4 * mm
            max_h = h - 4 * mm
            scale = min(max_w / iw, max_h / ih)

            draw_w = iw * scale
            draw_h = ih * scale

            img_x = x + (w - draw_w) / 2.0
            img_y = y + (h - draw_h) / 2.0

            c.drawImage(
                img,
                img_x,
                img_y,
                width=draw_w,
                height=draw_h,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass


def draw_material_program_block(c, x, y, w, h, material="", thickness="", program=""):
    draw_box(c, x, y, w, h)
    row_h = h / 3.0
    label_w = 26 * mm
    split_x = x + label_w

    c.line(split_x, y, split_x, y + h)
    c.line(x, y + row_h, x + w, y + row_h)
    c.line(x, y + 2 * row_h, x + w, y + 2 * row_h)

    labels = ["MATERIAL", "THICKNESS", "PROGRAM"]
    values = [str(material or ""), str(thickness or ""), str(program or "")]

    for i, (label, value) in enumerate(zip(labels, values)):
        cell_y = y + h - (i + 1) * row_h
        draw_cell_text(c, x, cell_y, label_w, row_h, label, align="left", font_size=10)
        draw_cell_text(c, split_x, cell_y, w - label_w, row_h, value, align="left", font_size=10)


def draw_doc_info_block(c, x, y, w, h, date_text="", dwg_no="", doc_no=""):
    draw_box(c, x, y, w, h)
    row_h = h / 3.0
    label_w = 24 * mm
    split_x = x + label_w

    c.line(split_x, y, split_x, y + h)
    c.line(x, y + row_h, x + w, y + row_h)
    c.line(x, y + 2 * row_h, x + w, y + 2 * row_h)

    labels = ["DATE", "DWG. NO.", "DOC. NO."]
    values = [str(date_text or ""), str(dwg_no or ""), str(doc_no or "")]

    for i, (label, value) in enumerate(zip(labels, values)):
        cell_y = y + h - (i + 1) * row_h
        draw_cell_text(c, x, cell_y, label_w, row_h, label, align="left", font_size=10)
        draw_cell_text(c, split_x, cell_y, w - label_w, row_h, value, align="left", font_size=10)


def draw_footer_note(c, page_w):
    c.setFont("Helvetica", 7)
    footer = "This document is generated automatically by PE Bag Drawing Generator."
    draw_text_center(c, page_w / 2, 7 * mm, footer, "Helvetica", 7)


def draw_main_frame(c, page_w, page_h, margin=8 * mm):
    c.setLineWidth(1.5)
    draw_box(c, margin, margin + 14 * mm, page_w - 2 * margin, page_h - 2 * margin - 22 * mm)


def draw_vertical_note_with_leader(c, text_lines, label_x, label_y, target_x, target_y, line_gap=5):
    dim_color = colors.HexColor("#165A8A")
    c.setStrokeColor(dim_color)
    c.setFillColor(colors.black)
    c.setLineWidth(1.0)

    leader_start_x = label_x + 8 * mm
    leader_start_y = label_y + 2 * mm
    c.line(leader_start_x, leader_start_y, target_x, target_y)

    c.saveState()
    c.translate(label_x, label_y)
    c.rotate(90)
    c.setFont("Helvetica", 11)
    current_y = 0
    for line in text_lines:
        c.drawString(0, current_y, line)
        current_y -= line_gap * mm / 3.0
    c.restoreState()


def draw_open_bag_in_area(
    c,
    area_x,
    area_y,
    area_w,
    area_h,
    bag_length_mm,
    bag_width_mm,
    seal_width_mm,
    distance_from_seal_mm,
    bag_fill,
    anti_static_text="Required",
    cleanroom_grade_text="Yes",
    remark_color_text="",
):
    extra_left = 42 * mm
    extra_right = 24 * mm
    extra_top = 20 * mm
    extra_bottom = 18 * mm

    usable_w = area_w - extra_left - extra_right
    usable_h = area_h - extra_top - extra_bottom

    scale_x = usable_w / (bag_width_mm * mm)
    scale_y = usable_h / (bag_length_mm * mm)
    scale = min(scale_x, scale_y)
    scale = min(scale, 2.4)
    scale = max(scale, 0.12)

    bag_draw_w = bag_width_mm * mm * scale
    bag_draw_h = bag_length_mm * mm * scale

    bag_x = area_x + extra_left + (usable_w - bag_draw_w) / 2
    bag_y = area_y + extra_bottom + (usable_h - bag_draw_h) / 2

    ellipse_h = min(6 * mm, max(2.8 * mm, bag_draw_w * 0.06))

    top_y = bag_y + bag_draw_h
    ellipse_bottom_y = top_y - ellipse_h / 2
    ellipse_top_y = top_y + ellipse_h / 2

    c.setFillColor(bag_fill)
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.2)
    c.rect(bag_x, bag_y, bag_draw_w, ellipse_bottom_y - bag_y, stroke=0, fill=1)

    path = c.beginPath()
    path.moveTo(bag_x, ellipse_bottom_y)
    path.curveTo(
        bag_x, ellipse_top_y,
        bag_x + bag_draw_w, ellipse_top_y,
        bag_x + bag_draw_w, ellipse_bottom_y
    )
    path.lineTo(bag_x + bag_draw_w, ellipse_bottom_y)
    path.lineTo(bag_x, ellipse_bottom_y)
    path.close()
    c.drawPath(path, stroke=0, fill=1)

    c.setStrokeColor(colors.black)
    c.setLineWidth(1.2)
    c.line(bag_x, bag_y, bag_x, ellipse_bottom_y)
    c.line(bag_x + bag_draw_w, bag_y, bag_x + bag_draw_w, ellipse_bottom_y)
    c.line(bag_x, bag_y, bag_x + bag_draw_w, bag_y)
    c.ellipse(
        bag_x,
        ellipse_bottom_y,
        bag_x + bag_draw_w,
        ellipse_top_y,
        stroke=1,
        fill=0,
    )

    seal_color = colors.HexColor("#C55A11")
    c.setStrokeColor(seal_color)
    seal_line_w = max(1.4, seal_width_mm * scale * 0.6)
    c.setLineWidth(seal_line_w)

    seal_center_y = bag_y + distance_from_seal_mm * mm * scale
    seal_center_y = max(bag_y + 4 * mm, min(seal_center_y, bag_y + bag_draw_h - 8 * mm))

    band_gap = max(2.0, seal_line_w * 1.3)
    y1 = seal_center_y - band_gap
    y2 = seal_center_y
    y3 = seal_center_y + band_gap

    c.line(bag_x, y1, bag_x + bag_draw_w, y1)
    c.line(bag_x, y2, bag_x + bag_draw_w, y2)
    c.line(bag_x, y3, bag_x + bag_draw_w, y3)

    dim_color = colors.HexColor("#165A8A")
    c.setStrokeColor(dim_color)
    c.setFillColor(dim_color)
    c.setLineWidth(1)

    top_dim_y = ellipse_top_y + 6 * mm
    c.line(bag_x, ellipse_top_y, bag_x, top_dim_y - 2 * mm)
    c.line(bag_x + bag_draw_w, ellipse_top_y, bag_x + bag_draw_w, top_dim_y - 2 * mm)

    draw_dimension_line(
        c,
        bag_x,
        top_dim_y,
        bag_x + bag_draw_w,
        top_dim_y,
        f"{bag_width_mm:g}mm +/-5mm",
        text_offset=-2 * mm,
        vertical=False,
    )

    right_dim_x = bag_x + bag_draw_w + 12 * mm
    c.setStrokeColor(dim_color)
    c.line(bag_x + bag_draw_w, bag_y, right_dim_x, bag_y)
    c.line(bag_x + bag_draw_w, ellipse_top_y, right_dim_x, ellipse_top_y)

    draw_dimension_line(
        c,
        right_dim_x,
        bag_y,
        right_dim_x,
        ellipse_top_y,
        f"{bag_length_mm:g}mm +/-5mm",
        text_offset=4 * mm,
        vertical=True,
    )

    left_dim_x = bag_x - 8 * mm
    c.line(bag_x, bag_y, left_dim_x, bag_y)
    c.line(bag_x, seal_center_y, left_dim_x, seal_center_y)

    draw_dimension_line(
        c,
        left_dim_x,
        bag_y,
        left_dim_x,
        seal_center_y,
        f"{distance_from_seal_mm:g}mm +/-1mm",
        text_offset=-5 * mm,
        vertical=True,
    )

    note_x = bag_x - 20 * mm
    note_y = seal_center_y + 40 * mm
    target_x = bag_x + bag_draw_w * 0.18
    target_y = y2

    draw_vertical_note_with_leader(
        c,
        text_lines=["Seal width", f"= {seal_width_mm:g}mm +/-1mm"],
        label_x=note_x,
        label_y=note_y,
        target_x=target_x,
        target_y=target_y,
        line_gap=10,
    )

    remark_x = area_x + 4 * mm
    remark_top_y = area_y + area_h - 10 * mm

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 11)
    c.drawString(remark_x, remark_top_y, "Remark")
    c.drawString(remark_x, remark_top_y - 7 * mm, f"Anti-Static : {anti_static_text}")
    c.drawString(remark_x, remark_top_y - 14 * mm, f"Cleanroom Grade : {cleanroom_grade_text}")
    c.drawString(remark_x, remark_top_y - 21 * mm, f"Color : {remark_color_text}")


def draw_pe_bag_proposal_pdf(output_path: str, record: dict):
    material = clean_str(record.get("material", ""))
    thickness = clean_str(record.get("thickness", ""))
    program = clean_str(record.get("program", ""))

    bag_length_mm = to_float(record["length"], "Length")
    bag_width_mm = to_float(record["width"], "Width")
    seal_width_mm = to_float(record["seal width"], "Seal Width")
    distance_from_seal_mm = to_float(record["distance from seal"], "Distance From Seal")

    date_text = clean_str(record.get("date", ""))
    dwg_no = clean_str(record.get("dwg no", ""))
    doc_no = clean_str(record.get("doc no", ""))

    anti_static_text = clean_str(record.get("anti-static", "Required"), "Required")
    cleanroom_grade_text = clean_str(record.get("cleanroom grade", "Yes"), "Yes")
    remark_color_text = clean_str(record.get("color", ""), "")

    hex_value = clean_str(record["hex"], "#FFF2CC")
    try:
        bag_fill = HexColor(hex_value)
    except Exception:
        bag_fill = HexColor("#FFF2CC")

    c = canvas.Canvas(output_path, pagesize=A4)
    page_w, page_h = A4

    draw_main_frame(c, page_w, page_h)

    left = 8 * mm
    right = page_w - 8 * mm
    frame_bottom = 22 * mm
    frame_top = page_h - 8 * mm
    frame_w = right - left

    title_y = frame_top - 8 * mm
    c.setFont("Helvetica-Bold", 18)
    draw_text_center(c, page_w / 2, title_y, "PE BAG DRAWING PROPOSAL", "Helvetica-Bold", 18)

    c.setLineWidth(1)
    c.line(left + 62 * mm, title_y - 2 * mm, right - 62 * mm, title_y - 2 * mm)
    c.line(left + 62 * mm, title_y - 3.5 * mm, right - 62 * mm, title_y - 3.5 * mm)

    rev_top = title_y - 8 * mm
    rev_x = left + 2 * mm
    rev_w = frame_w - 4 * mm
    draw_revision_table(c, rev_x, rev_top, rev_w, row_h=7 * mm, data_rows=4)

    bottom_blocks_h = 18 * mm
    drawing_top = rev_top - 5 * 7 * mm - 3 * mm
    drawing_y = frame_bottom + 2 * mm + bottom_blocks_h
    drawing_h = drawing_top - drawing_y
    drawing_x = left + 2 * mm
    drawing_w = frame_w - 4 * mm

    draw_box(c, drawing_x, drawing_y, drawing_w, drawing_h)

    draw_open_bag_in_area(
        c,
        drawing_x,
        drawing_y,
        drawing_w,
        drawing_h,
        bag_length_mm=bag_length_mm,
        bag_width_mm=bag_width_mm,
        seal_width_mm=seal_width_mm,
        distance_from_seal_mm=distance_from_seal_mm,
        bag_fill=bag_fill,
        anti_static_text=anti_static_text,
        cleanroom_grade_text=cleanroom_grade_text,
        remark_color_text=remark_color_text,
    )

    block_y = frame_bottom + 2 * mm
    block_h = 18 * mm

    approval_w = 42 * mm
    logo_w = 30 * mm
    material_w = 58 * mm
    doc_w = 60 * mm

    approval_x = left + 2 * mm
    logo_x = approval_x + approval_w
    material_x = logo_x + logo_w
    doc_x = material_x + material_w

    draw_approval_block(c, approval_x, block_y, approval_w, block_h)
    draw_logo_only_block(c, logo_x, block_y, logo_w, block_h, logo_path=str(get_base_dir() / LOGO_FILE))
    draw_material_program_block(
        c,
        material_x,
        block_y,
        material_w,
        block_h,
        material=material,
        thickness=thickness,
        program=program,
    )
    draw_doc_info_block(
        c,
        doc_x,
        block_y,
        doc_w,
        block_h,
        date_text=date_text,
        dwg_no=dwg_no,
        doc_no=doc_no,
    )

    draw_footer_note(c, page_w)

    c.showPage()
    c.save()


def main():
    base_dir = get_base_dir()

    input_path = base_dir / INPUT_FILE
    output_dir = base_dir / OUTPUT_DIR

    if not input_path.exists():
        show_error("Error", f"Excel file not found:\n{input_path}")
        return

    os.makedirs(output_dir, exist_ok=True)

    try:
        df = pd.read_excel(input_path, sheet_name=SHEET_NAME, header=1)
        df = normalize_columns(df)
        df = df.dropna(how="all")
    except Exception as e:
        show_error("Error", f"Failed to read Excel file.\n\n{e}")
        return

    if df.empty:
        show_error("Error", "No data rows found in the Excel sheet.")
        return

    created_files = []
    skipped_rows = []

    for idx, row in df.iterrows():
        row_dict = row.to_dict()
        program_name = safe_filename(row_dict.get("program", ""), fallback=f"Row_{idx+2}")
        filename = f"PE_Bag_Proposal_{idx+2}_{program_name}.pdf"
        output_path = output_dir / filename

        try:
            draw_pe_bag_proposal_pdf(str(output_path), row_dict)
            created_files.append(str(output_path))
        except Exception as e:
            skipped_rows.append(f"Row {idx+2}: {e}")

    if created_files and not skipped_rows:
        show_info("Done", f"Done, {len(created_files)} PDF(s) created.\n\nOutput folder:\n{output_dir}")
    elif created_files and skipped_rows:
        msg = (
            f"Done, {len(created_files)} PDF(s) created.\n\n"
            f"Some rows were skipped:\n"
            + "\n".join(skipped_rows[:10])
        )
        if len(skipped_rows) > 10:
            msg += f"\n... and {len(skipped_rows) - 10} more."
        msg += f"\n\nOutput folder:\n{output_dir}"
        show_info("Completed with Warnings", msg)
    else:
        show_error("Error", "No PDFs were created.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        detail = "".join(traceback.format_exception_only(type(e), e)).strip()
        show_error("Unexpected Error", detail)