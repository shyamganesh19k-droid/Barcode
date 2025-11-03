from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
import os, difflib
from reportlab.pdfgen import canvas

app = Flask(__name__)

EXCEL_FILE = "bom_data.xlsx"
OUTPUT_DIR = os.path.join(app.static_folder, "labels")
os.makedirs(OUTPUT_DIR, exist_ok=True)

required_cols = {
    "sku": ["sku", "item", "itemcode"],
    "bom_desc": ["bom description", "bomdesc", "description"],
    "bom_line": ["bom line description", "bomline", "line description"],
    "isbn": ["isbn"],
    "mrp": ["mrp"]
}

# Fonts
try:
    FONT_TITLE = ImageFont.truetype("arialbd.ttf", 40)
    FONT_HEADING = ImageFont.truetype("arialbd.ttf", 32)
    FONT_BODY = ImageFont.truetype("arial.ttf", 28)
    FONT_SMALL = ImageFont.truetype("arial.ttf", 28)
except:
    FONT_TITLE = FONT_HEADING = FONT_BODY = FONT_SMALL = ImageFont.load_default()


def load_data():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(), {}
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip().str.lower()
    col_map = {}
    for logical, options in required_cols.items():
        for opt in options:
            matches = difflib.get_close_matches(opt.lower(), df.columns, n=1, cutoff=0.6)
            if matches:
                col_map[logical] = matches[0]
                break
    return df, col_map


def generate_label(sku_input, override_mrp=None):
    df, col_map = load_data()
    if df.empty or "sku" not in col_map:
        return None, "Excel file or SKU column missing."

    sku_upper = sku_input.upper()
    bom_info = df[df[col_map["sku"]].astype(str).str.upper() == sku_upper]
    if bom_info.empty:
        return None, "❌ SKU not found!"

    desc = str(bom_info.iloc[0].get(col_map.get("bom_desc", ""), "N/A"))
    isbn = str(bom_info.iloc[0].get(col_map.get("isbn", ""), "N/A"))
    mrp = override_mrp if override_mrp else str(bom_info.iloc[0].get(col_map.get("mrp", ""), "N/A"))
    qty = len(bom_info)
    line_items = bom_info[col_map["bom_line"]].tolist() if "bom_line" in col_map else []

    # Barcode
    barcode_base = os.path.join(OUTPUT_DIR, f"{sku_input}_barcode")
    code128 = barcode.get("code128", sku_input, writer=ImageWriter())
    code128.save(barcode_base, {"write_text": False})
    barcode_file = barcode_base + ".png"
    barcode_img = Image.open(barcode_file).resize((420, 140))

    # Label canvas
    label = Image.new("RGB", (1200, 891), "white")
    draw = ImageDraw.Draw(label)

    draw.rectangle([10, 10, 1190, 881], outline="black", width=4)
    draw.text((30, 25), desc, font=FONT_TITLE, fill="black")

    y = 85
    draw.text((30, y), "CONTENTS OF KIT:", font=FONT_HEADING, fill="black")
    y += 40
    for i, item in enumerate(line_items[:12], start=1):
        draw.text((50, y), f"{i}. {item}", font=FONT_SMALL, fill="black")
        y += 42

    # Info box
    box_left, box_top, box_right, box_bottom = 20, 680, 560, 870
    draw.rectangle([box_left, box_top, box_right, box_bottom], outline="black", width=3)
    info_lines = [f"ISBN: {isbn}", f"Qty: {qty} Items", f"MRP: ₹{mrp}"]
    total_text_height = len(info_lines) * 38
    start_y = box_top + (190 - total_text_height) // 2
    for i, line in enumerate(info_lines):
        draw.text((box_left + 20, start_y + i * 42), line, font=FONT_BODY, fill="black")

    # Barcode box
    draw.rectangle([580, 680, 1180, 870], outline="black", width=3)
    label.paste(barcode_img, (670, 690))
    draw.text((800, 830), sku_input, font=FONT_BODY, fill="black")

    # Save
    label_filename = f"{sku_input}_label.png"
    label.save(os.path.join(OUTPUT_DIR, label_filename))
    return label_filename, None


@app.route("/", methods=["GET", "POST"])
def index():
    label = None
    error = None
    sku_input = None
    product_name = None
    mrp = None

    if request.method == "POST":
        sku_input = (request.form.get("sku") or "").strip()
        override_mrp = request.form.get("mrp") or None
        df, col_map = load_data()

        if not sku_input:
            error = "Please enter a SKU."
        elif df.empty or "sku" not in col_map:
            error = "Excel file missing or invalid format."
        else:
            bom_info = df[df[col_map["sku"]].astype(str).str.upper() == sku_input.upper()]
            if not bom_info.empty:
                product_name = str(bom_info.iloc[0].get(col_map.get("bom_desc", ""), "N/A"))
                mrp = override_mrp if override_mrp else str(bom_info.iloc[0].get(col_map.get("mrp", ""), "N/A"))
            else:
                error = "❌ SKU not found!"

        if "generate" in request.form and not error:
            label, error = generate_label(sku_input, override_mrp)

    return render_template("index.html",
                           label=label,
                           error=error,
                           sku=sku_input,
                           product_name=product_name,
                           mrp=mrp)

@app.route("/download/<sku>/<mrp>")
def download_pdf(sku, mrp):
    mrp_val = None if mrp == "NA" else mrp
    label_filename, error = generate_label(sku, mrp_val)
    if not label_filename:
        return f"Error: {error}"

    label_path = os.path.join(OUTPUT_DIR, label_filename)
    pdf_path = os.path.join(OUTPUT_DIR, f"{sku}_label.pdf")
    c = canvas.Canvas(pdf_path, pagesize=(288, 214))
    c.drawImage(label_path, 0, 0, width=288, height=214)
    c.save()

    return send_file(pdf_path, as_attachment=True)


@app.route("/cancel")
def cancel():
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
