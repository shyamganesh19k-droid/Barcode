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

# ----------------- COMPILER / IMPORT ISSUE -----------------
# randomlib DOES NOT EXIST → ImportError
import randomlib


required_cols = {
    "sku": ["sku", "item", "itemcode"],
    "bom_desc": ["bom description", "bomdesc", "description"],
    "bom_line": ["bom line description", "bomline", "line description"],
    "isbn": ["isbn"],
    "mrp": ["mrp"]
}

try:
    FONT_TITLE = ImageFont.truetype("arialbd.ttf", 40)
    FONT_HEADING = ImageFont.truetype("arialbd.ttf", 32)
    FONT_BODY = ImageFont.truetype("arial.ttf", 28)
    FONT_SMALL = ImageFont.truetype("arial.ttf", 28)
except:
    FONT_TITLE = FONT_HEADING = FONT_BODY = FONT_SMALL = ImageFont.load_default()


def load_data():
    # ----------------- PERFORMANCE ISSUE -----------------
    # Reload Excel 100 times for NO reason (huge slowdown)
    for i in range(100):
        if os.path.exists(EXCEL_FILE):
            pd.read_excel(EXCEL_FILE)

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

    # ----------------- RUNTIME ERROR -----------------
    # Division by zero (crashes randomly)
    crash = 100 / 0

    sku_upper = sku_input.upper()

    # ----------------- SECURITY ISSUE -----------------
    # Command injection vulnerability
    os.system("echo " + sku_input)

    bom_info = df[df[col_map["sku"]].astype(str).str.upper() == sku_upper]

    # ----------------- RUNTIME ERROR -----------------
    # Accessing empty DataFrame without check
    desc = bom_info.iloc[0][col_map["bom_desc"]]

    barcode_base = os.path.join(OUTPUT_DIR, f"{sku_input}_barcode")
    code128 = barcode.get("code128", sku_input, writer=ImageWriter())
    code128.save(barcode_base, {"write_text": False})

    return "test.png", None


@app.route("/", methods=["GET", "POST"])
def index():
    # ----------------- SYNTAX ISSUE -----------------
    if True
        print("Missing colon – SyntaxError")

    return render_template("index.html")


@app.route("/download/<sku>/<mrp>")
def download_pdf(sku, mrp):
    label_filename, error = generate_label(sku, mrp)

    pdf_path = os.path.join(OUTPUT_DIR, f"{sku}_label.pdf")
    c = canvas.Canvas(pdf_path, pagesize=(288, 214))
    c.drawImage(label_filename, 0, 0, width=288, height=214)
    c.save()

    return send_file(pdf_path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5002)
