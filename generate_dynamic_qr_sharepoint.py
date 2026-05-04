"""
generate_dynamic_qr_sharepoint.py (tight content-hugging rounded border + no overlap)
"""

import os
import math
import logging
import sys
from typing import Optional

import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont

# ---------------------- CONFIG ----------------------
EXCEL_PATH = r"C:\QRAssets\Dynamic_Asset_List.xlsx"
LOGO_PATH = r"C:\QRAssets\company_logo.png"
OUTPUT_FOLDER = "C:/QRAssets/qr_codes"
PDF_OUTPUT = os.path.join(OUTPUT_FOLDER, "print_pages.pdf")

BASE_SHAREPOINT_DISPFORM = "https://everrenew.sharepoint.com/sites/ITassetmanagement/Lists/HardwareAssetList/DispForm.aspx?ID="

# Print/A4 settings (300 DPI)
DPI = 300
A4_WIDTH_PX = int(8.27 * DPI)
A4_HEIGHT_PX = int(11.69 * DPI)

COLS = 3
ROWS = 6
STICKERS_PER_PAGE = COLS * ROWS

PAGE_MARGIN = 120
GUTTER_X = 60
GUTTER_Y = 40

# QR and sticker style
QR_BOX_SIZE = 6
QR_BORDER = 2
QR_LOGO_RATIO = 0.20
RIGHT_LOGO_MAX_HEIGHT = 180
BACKGROUND_COLOR = "white"

# border around content (will be drawn tightly around QR+text+logo)
BORDER_COLOR = "#000000"
BORDER_THICKNESS = 6
BORDER_RADIUS = 24
CONTENT_PADDING = 10  # padding between content and border (tight)

# Fonts
try:
    FONT_PATH = "arial.ttf"
    default_font = ImageFont.truetype(FONT_PATH, 40)
except Exception:
    default_font = ImageFont.load_default()

# Ensure output dir
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s",
                    handlers=[logging.StreamHandler(sys.stdout)])
log = logging.getLogger("dynamic_qr")

# ---------------------- HELPERS ----------------------
def safe_filename(s: str) -> str:
    return "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in s)[:100]

def build_item_link(row: pd.Series) -> Optional[str]:
    if "ItemLink" in row and pd.notna(row["ItemLink"]) and str(row["ItemLink"]).strip() != "":
        return str(row["ItemLink"]).strip()
    if "ID" in row and pd.notna(row["ID"]):
        return BASE_SHAREPOINT_DISPFORM + str(int(row["ID"]))
    return None

def generate_qr_image(data: str, logo_path: Optional[str] = None,
                      qr_box_size: int = QR_BOX_SIZE, border: int = QR_BORDER) -> Image.Image:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=qr_box_size,
        border=border
    )
    qr.add_data(data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")

    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image.open(logo_path).convert("RGBA")
            qr_w, qr_h = qr_img.size
            logo_size = max(24, int(qr_w * QR_LOGO_RATIO))
            if hasattr(Image, "Resampling"):
                logo.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)
            else:
                logo.thumbnail((logo_size, logo_size), Image.ANTIALIAS)
            lx, ly = logo.size
            pos = ((qr_w - lx) // 2, (qr_h - ly) // 2)
            qr_img.paste(logo, pos, mask=logo)
        except Exception as e:
            log.warning(f"Could not embed logo inside QR: {e}")
    return qr_img.convert("RGB")

# ---------------------- STICKER LAYOUT ----------------------
usable_width = A4_WIDTH_PX - 2 * PAGE_MARGIN - (COLS - 1) * GUTTER_X
usable_height = A4_HEIGHT_PX - 2 * PAGE_MARGIN - (ROWS - 1) * GUTTER_Y
sticker_w = usable_width // COLS
sticker_h = usable_height // ROWS

# QR target size = 1 inch (≈300 px)
qr_target_size = 300
right_section_w = sticker_w - qr_target_size - 30
if right_section_w < 150:
    right_section_w = max(150, sticker_w // 3)

log.info(f"A4: {A4_WIDTH_PX}x{A4_HEIGHT_PX} | Sticker: {sticker_w}x{sticker_h} | QR size: {qr_target_size}px")

# ---------------------- MAIN ----------------------
def main():
    log.info("Loading Excel...")
    if not os.path.exists(EXCEL_PATH):
        log.error(f"Excel file not found: {EXCEL_PATH}")
        return

    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    log.info(f"Rows loaded: {len(df)}")

    individual_folder = os.path.join(OUTPUT_FOLDER, "individual")
    os.makedirs(individual_folder, exist_ok=True)

    page_images = []
    stickers_on_current_page = 0
    current_page = Image.new("RGB", (A4_WIDTH_PX, A4_HEIGHT_PX), BACKGROUND_COLOR)

    for idx, row in df.iterrows():
        try:
            asset_id = str(row.get("Title", f"asset_{idx}"))
            brand = str(row.get("Brand", ""))
            serial = str(row.get("Serial number", ""))
            warranty = str(row.get("Warranty Expiry", ""))
            item_link = build_item_link(row)
            if not item_link:
                log.warning(f"Skipping {asset_id} — no ItemLink or ID found.")
                continue

            safe_name = safe_filename(asset_id)
            individual_path = os.path.join(individual_folder, f"{safe_name}.png")
            if os.path.exists(individual_path):
                log.info(f"Skipping existing sticker for {asset_id}")
                continue

            # Generate QR and resize
            qr_img = generate_qr_image(item_link, logo_path=LOGO_PATH)
            qr_img = qr_img.resize((qr_target_size, qr_target_size),
                                   Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.ANTIALIAS)

            # Create sticker canvas
            sticker = Image.new("RGB", (sticker_w, sticker_h), BACKGROUND_COLOR)
            draw = ImageDraw.Draw(sticker)

            # Layout spacing (tight)
            left_margin = 18
            right_block_x = left_margin + qr_target_size + 12  # right block start
            padding_inside = 6  # later used for border padding around content

            # Place QR centered vertically but a bit up to leave space
            qr_x = left_margin
            qr_y = max(8, (sticker_h - qr_target_size) // 2 - 8)  # slight upward nudge
            sticker.paste(qr_img, (qr_x, qr_y))

            # Prepare right content (Asset ID text and logo), measure then center vertically
            # Font selection and dynamic size to fit width
            font_size = 42
            try:
                font = ImageFont.truetype(FONT_PATH, font_size)
            except Exception:
                font = default_font

            text = f"Asset ID: {asset_id}"

            # Reduce font size if it doesn't fit the right area
            max_text_width = sticker_w - right_block_x - 18
            while True:
                bbox = draw.textbbox((0, 0), text, font=font)
                text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
                if text_w <= max_text_width or font_size <= 10:
                    break
                font_size -= 2
                try:
                    font = ImageFont.truetype(FONT_PATH, font_size)
                except Exception:
                    font = default_font
                    break

            # Load & scale right-side logo
            right_logo = None
            right_logo_w = right_logo_h = 0
            if os.path.exists(LOGO_PATH):
                right_logo = Image.open(LOGO_PATH).convert("RGBA")
                max_logo_w = sticker_w - right_block_x - 18
                max_logo_h = RIGHT_LOGO_MAX_HEIGHT
                if hasattr(Image, "Resampling"):
                    right_logo.thumbnail((max_logo_w, max_logo_h), Image.Resampling.LANCZOS)
                else:
                    right_logo.thumbnail((max_logo_w, max_logo_h), Image.ANTIALIAS)
                right_logo_w, right_logo_h = right_logo.size

            spacing = 10
            total_right_h = text_h + (spacing + right_logo_h if right_logo is not None else 0)
            # center the right block vertically aligned with QR
            start_y_block = qr_y + (qr_target_size - total_right_h) // 2

            # Draw text
            text_x = right_block_x + ( (sticker_w - right_block_x - 18) - text_w ) // 2
            text_y = start_y_block
            draw.text((text_x, text_y), text, fill="black", font=font)

            # Paste logo below text
            logo_x = logo_y = None
            if right_logo is not None:
                logo_x = right_block_x + ((sticker_w - right_block_x - 18) - right_logo_w) // 2
                logo_y = text_y + text_h + spacing
                sticker.paste(right_logo, (logo_x, logo_y), mask=right_logo)

            # === compute tight content bounding box ===
            content_left = qr_x
            content_top = min(qr_y, text_y)
            content_right = qr_x + qr_target_size
            # extend right to include text and logo if wider
            content_right = max(content_right, text_x + text_w)
            if right_logo is not None:
                content_right = max(content_right, logo_x + right_logo_w)
            content_bottom = qr_y + qr_target_size
            if right_logo is not None:
                content_bottom = max(content_bottom, logo_y + right_logo_h)

            # add small padding and clamp inside sticker
            pad = CONTENT_PADDING
            rect_left = max(2, content_left - pad)
            rect_top = max(2, content_top - pad)
            rect_right = min(sticker_w - 2, content_right + pad)
            rect_bottom = min(sticker_h - 2, content_bottom + pad)

            # Draw tight rounded border around content
            try:
                draw.rounded_rectangle(
                    [(rect_left, rect_top), (rect_right, rect_bottom)],
                    radius=BORDER_RADIUS,
                    outline=BORDER_COLOR,
                    width=BORDER_THICKNESS
                )
            except Exception:
                # fallback to normal rectangle if rounded not available
                draw.rectangle([(rect_left, rect_top), (rect_right, rect_bottom)],
                               outline=BORDER_COLOR, width=BORDER_THICKNESS)

            # Save individual sticker
            sticker.save(individual_path)
            log.info(f"Saved sticker: {individual_path}")

            # Place sticker onto page canvas
            pos_in_page = stickers_on_current_page
            col = pos_in_page % COLS
            rowpos = pos_in_page // COLS
            paste_x = PAGE_MARGIN + col * (sticker_w + GUTTER_X)
            paste_y = PAGE_MARGIN + rowpos * (sticker_h + GUTTER_Y)
            current_page.paste(sticker, (paste_x, paste_y))
            stickers_on_current_page += 1

            # if full page, append and create new page
            if stickers_on_current_page >= STICKERS_PER_PAGE:
                page_images.append(current_page)
                current_page = Image.new("RGB", (A4_WIDTH_PX, A4_HEIGHT_PX), BACKGROUND_COLOR)
                stickers_on_current_page = 0

        except Exception as e:
            log.exception(f"Failed to process row {idx}: {e}")
            continue

    # append last page if any
    if stickers_on_current_page > 0:
        page_images.append(current_page)

    if not page_images:
        log.info("No new stickers generated.")
        return

    # save PDF pages
    try:
        page_images[0].save(PDF_OUTPUT, "PDF", resolution=DPI, save_all=True, append_images=page_images[1:])
        log.info(f"Saved PDF: {PDF_OUTPUT} (pages: {len(page_images)})")
    except Exception as e:
        log.exception(f"Failed to save PDF: {e}")

    log.info("Done.")

if __name__ == "__main__":
    main()
