import io
import os
import re
from pathlib import Path
from typing import Dict, List, Tuple

from flask import Flask, request, send_file, render_template_string
from pypdf import PdfReader
from PIL import Image as PILImage
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

APP_DIR = Path(__file__).resolve().parent
HEADER_IMG = APP_DIR / "yancey_cat_header_strip.png"
app = Flask(__name__)

HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Convert to CVA Format</title>
  <style>
    body { font-family: Arial, sans-serif; background:#f5f5f5; margin:0; padding:40px; }
    .wrap { max-width:760px; margin:0 auto; background:white; padding:32px; border-radius:16px; box-shadow:0 8px 30px rgba(0,0,0,.08); }
    h1 { margin-top:0; }
    .hint { color:#555; line-height:1.5; }
    .box { border:2px dashed #bbb; border-radius:12px; padding:28px; margin:24px 0; text-align:center; background:#fafafa; }
    input[type=file] { margin:10px 0 20px 0; }
    button { background:#111; color:white; border:none; border-radius:10px; padding:12px 18px; font-size:16px; cursor:pointer; }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>Convert to CVA Format</h1>
    <p class="hint">Upload a Yancey PMAP quote PDF and download the CVA version.</p>
    <form method="post" enctype="multipart/form-data">
      <div class="box">
        <input type="file" name="file" accept=".pdf" required />
        <br />
        <button type="submit">Convert to CVA Format</button>
      </div>
    </form>
  </div>
</body>
</html>
"""

def extract_text_from_bytes(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    return "\n".join((page.extract_text() or "") for page in reader.pages)

def find_first(pattern: str, text: str, default: str = "") -> str:
    m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
    return m.group(1).strip() if m else default

def parse_equipment(text: str) -> Dict[str, str]:
    labels = [
        "Make", "Model", "Serial Number or Range", "Start Hours",
        "Travel Zone", "Service Interval", "Agreement Term", "Agreement Usage"
    ]
    out = {k: "" for k in labels}
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for idx, line in enumerate(lines):
        if line in out and idx > 0:
            out[line] = lines[idx - 1]
    return out

def parse_services_total_column(text: str) -> List[Tuple[str, str, str]]:
    names = ["Initial Service", "A Service", "B Service", "C Service", "D Service"]
    pattern = re.compile(
        r"(Initial Service|A Service|B Service|C Service|D Service)\s+"
        r"([0-9,]+\s+hours\s*/\s*[0-9]+\s+months)"
        r"(.*?)(?=(?:Initial Service|A Service|B Service|C Service|D Service|Additional Charges))",
        re.DOTALL,
    )
    rows = []
    for m in pattern.finditer(text):
        name = m.group(1).strip()
        interval = m.group(2).strip()
        tail = m.group(3)
        amounts = re.findall(r"\$\d[\d,]*\.\d{2}", tail)
        total = amounts[3] if len(amounts) >= 4 else (amounts[-1] if amounts else "")
        rows.append((name, interval, total))

    ordered = []
    seen = set()
    for name in names:
        for row in rows:
            if row[0] == name and name not in seen:
                ordered.append(row)
                seen.add(name)
                break
    return ordered

def parse_quote(pdf_bytes: bytes) -> Dict:
    text = extract_text_from_bytes(pdf_bytes)
    return {
        "quote_id": find_first(r"Quote ID:?\s*([0-9]+)", text, ""),
        "cost_per_hour": find_first(r"Cost Per Hour\s*\$?([\d,]+\.\d{2})", text, ""),
        "services": parse_services_total_column(text),
        "equipment": parse_equipment(text),
    }

def output_filename(data):
    import re

    def clean(value):
        value = (value or "").strip()
        value = re.sub(r"\s+", "", value)
        value = re.sub(r"[^A-Za-z0-9]", "", value)
        return value

    eq = data.get("equipment", {})
    serial_range = eq.get("Serial Number or Range", "")

    # Take FIRST serial from range
    if "-" in serial_range:
        serial = serial_range.split("-")[0]
    else:
        serial = serial_range

    serial = clean(serial)

    if not serial:
        return "CVA.pdf"

    return f"{serial}_CVA.pdf"
  
def build_pdf_bytes(data: Dict) -> bytes:
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, leftMargin=22, rightMargin=22, topMargin=12, bottomMargin=22)
    styles = getSampleStyleSheet()
    body = []

    pil = PILImage.open(HEADER_IMG)
    header = RLImage(str(HEADER_IMG))
    header.drawWidth = doc.width
    header.drawHeight = doc.width * (pil.size[1] / pil.size[0])
    body.append(header)
    body.append(Spacer(1, 10))

    small_label = ParagraphStyle("small_label", parent=styles["Normal"], fontSize=7, textColor=colors.HexColor("#333333"))
    value_style = ParagraphStyle("value_style", parent=styles["Normal"], fontSize=10, leading=11)
    service_style = ParagraphStyle("service_style", parent=styles["Normal"], fontSize=10, leading=12)
    amount_style = ParagraphStyle("amount_style", parent=styles["Normal"], fontSize=10, alignment=2)
    total_hdr = ParagraphStyle("total_hdr", parent=styles["Normal"], alignment=1, fontSize=12)

    eq = data["equipment"]
    eq_rows = [
        [Paragraph("Make", small_label), Paragraph("Model", small_label), Paragraph("Serial Number or Range", small_label), Paragraph("Start Hours", small_label)],
        [Paragraph(eq["Make"], value_style), Paragraph(eq["Model"], value_style), Paragraph(eq["Serial Number or Range"], value_style), Paragraph(eq["Start Hours"], value_style)],
        [Paragraph("Travel Zone", small_label), Paragraph("Service Interval", small_label), Paragraph("Agreement Term", small_label), Paragraph("Agreement Usage", small_label)],
        [Paragraph(eq["Travel Zone"], value_style), Paragraph(eq["Service Interval"], value_style), Paragraph(eq["Agreement Term"], value_style), Paragraph(eq["Agreement Usage"], value_style)],
    ]
    eq_table = Table(eq_rows, colWidths=[doc.width / 4] * 4)
    eq_table.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 1), (-1, 1), 0.8, colors.HexColor("#7f7f7f")),
        ("INNERGRID", (0, 1), (-1, 1), 0.8, colors.HexColor("#7f7f7f")),
        ("BOX", (0, 3), (-1, 3), 0.8, colors.HexColor("#7f7f7f")),
        ("INNERGRID", (0, 3), (-1, 3), 0.8, colors.HexColor("#7f7f7f")),
    ]))
    body.append(eq_table)
    body.append(Spacer(1, 16))

    body.append(Paragraph("<b>Total</b>", total_hdr))
    body.append(Spacer(1, 8))

    service_rows = []
    for name, interval, total in data["services"]:
        service_rows.append([
            Paragraph(f"{name}<br/><font size='8'>{interval}</font>", service_style),
            Paragraph(f"<b>{total}</b>", amount_style),
        ])
    st = Table(service_rows, colWidths=[doc.width * 0.74, doc.width * 0.26])
    st.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.35, colors.HexColor("#8f8f8f")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    body.append(st)
    body.append(Spacer(1, 16))

    body.append(Paragraph("Cost Per Hour", service_style))
    cost_box = Table([[f"${data['cost_per_hour']}"]], colWidths=[120], rowHeights=[38])
    cost_box.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.25, colors.HexColor("#d0a100")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 13),
    ]))
    body.append(Spacer(1, 4))
    body.append(cost_box)

    doc.build(body)
    output.seek(0)
    return output.getvalue()

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML)

    file = request.files.get("file")
    if not file or not file.filename.lower().endswith(".pdf"):
        return "Please upload a PDF file.", 400

    data = parse_quote(file.read())
    converted = build_pdf_bytes(data)
    out_name = output_filename(data)

    return send_file(io.BytesIO(converted), mimetype="application/pdf", as_attachment=True, download_name=out_name)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
