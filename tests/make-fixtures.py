"""Creates the synthetic test documents in tests/fixtures (no real data).

Requires python-docx, openpyxl and LibreOffice (for the PDF):
    python3 tests/make-fixtures.py
"""
import datetime
import os
import subprocess

from docx import Document
from docx.enum.text import WD_BREAK
from openpyxl import Workbook
from openpyxl.comments import Comment

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures")
os.makedirs(OUT, exist_ok=True)


def make_docx(path):
    d = Document()
    d.add_heading("Environmental Impact Assessment - Test Report", 0)
    d.add_heading("1 Introduction", 1)
    d.add_paragraph("This synthetic report checks headings, lists and merged table cells. "
                    "The predicted noise level is 55.5 dB(A) at receptor R-03.")
    d.add_heading("1.1 Scope", 2)
    for t in ["Air quality monitoring", "Noise and vibration", "Marine water quality"]:
        d.add_paragraph(t, style="List Bullet")
    for t in ["Baseline survey", "Impact assessment", "Mitigation"]:
        d.add_paragraph(t, style="List Number")
    p = d.add_paragraph("Bold text: ")
    p.add_run("critical receptor").bold = True
    p.add_run(" and ")
    p.add_run("italic note").italic = True
    p.add_run(".")

    d.add_heading("2 Results", 1)
    d.add_paragraph("Table 2-1 summarises the ambient air results.")
    t = d.add_table(rows=5, cols=4)
    t.style = "Table Grid"
    hdr = ["Parameter", "Location", "Result (µg/m³)", "Limit (µg/m³)"]
    for i, h in enumerate(hdr):
        t.cell(0, i).text = h
    data = [
        ["PM10", "AQ-01", "48.2", "150"],
        ["PM10", "AQ-02", "61.7", "150"],
        ["NO2", "AQ-01", "21.4", "200"],
        ["NO2", "AQ-02", "19.9", "200"],
    ]
    for r, row in enumerate(data, start=1):
        for c, v in enumerate(row):
            t.cell(r, c).text = v
    # vertical merges: parameter column; horizontal merge: limit column for NO2
    a = t.cell(1, 0).merge(t.cell(2, 0)); a.text = "PM10"
    b = t.cell(3, 0).merge(t.cell(4, 0)); b.text = "NO2"
    c = t.cell(3, 3).merge(t.cell(4, 3)); c.text = "200"

    d.add_paragraph("")
    t2 = d.add_table(rows=3, cols=3)
    t2.style = "Table Grid"
    m = t2.cell(0, 0).merge(t2.cell(0, 2)); m.text = "Section A: General comments"
    t2.cell(1, 0).text = "A.1"; t2.cell(1, 1).text = "Clarify the fill source."; t2.cell(1, 2).text = "Open"
    t2.cell(2, 0).text = "A.2"; t2.cell(2, 1).text = "Provide groundwater data, 12 samples."; t2.cell(2, 2).text = "Closed"

    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_heading("3 Conclusions", 1)
    d.add_paragraph("All 4 monitored parameters complied with the limits. Total cost was 1,250.75 AED.")
    sec = d.sections[0]
    sec.header.paragraphs[0].text = "TEST REPORT HEADER - Confidential"
    sec.footer.paragraphs[0].text = "Footer text 2026"
    d.save(path)


def make_xlsx(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Monitoring"
    ws["A1"] = "Ambient Noise Monitoring Results"
    ws.merge_cells("A1:E1")
    ws.append(["Station", "Period", "LAeq dB(A)", "Limit dB(A)", "Date"])
    rows = [
        ["N-01", "Day", 55.46, 65, datetime.date(2026, 3, 14)],
        ["N-01", "Night", 47.1, 55, datetime.date(2026, 3, 14)],
        ["N-02", "Day", 61.234, 65, datetime.date(2026, 3, 15)],
        ["N-02", "Night", 52, 55, datetime.date(2026, 3, 15)],
    ]
    for r in rows:
        ws.append(r)
    ws.merge_cells("A3:A4")
    ws.merge_cells("A5:A6")
    for r in range(3, 7):
        ws.cell(r, 3).number_format = "0.0"
        ws.cell(r, 5).number_format = "yyyy-mm-dd"
    ws["A9"] = "Exceedance rate"
    ws["B9"] = 0.125
    ws["B9"].number_format = "0.0%"
    ws["A10"] = "Budget (AED)"
    ws["B10"] = 1234567.891
    ws["B10"].number_format = "#,##0.00"
    ws["A11"] = "Average day level"
    ws["B11"] = "=AVERAGE(C3,C5)"
    ws["A9"].comment = Comment("Based on 8 measurements", "Reviewer")
    ws.row_dimensions[12].hidden = True
    ws["A12"] = "Hidden row value"

    ws2 = wb.create_sheet("Hidden data")
    ws2["A1"] = "Secret calibration factor"
    ws2.sheet_state = "hidden"

    ws3 = wb.create_sheet("Water Quality")
    ws3.append(["Parameter", "Unit", "WQ-01", "WQ-02"])
    ws3.append(["pH", "-", 8.1, 8.2])
    ws3.append(["TSS", "mg/L", 12, 15.5])
    wb.save(path)


make_docx(os.path.join(OUT, "synthetic-report.docx"))
make_xlsx(os.path.join(OUT, "synthetic-monitoring.xlsx"))
# PDF via LibreOffice (an untagged PDF: exercises the layout-only path)
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", OUT,
                os.path.join(OUT, "synthetic-report.docx")], check=True, capture_output=True)
# an untagged PDF of the same document: exercises the layout-only path
untagged = 'pdf:writer_pdf_Export:{"UseTaggedPDF":{"type":"boolean","value":"false"}}'
tmp = os.path.join(OUT, "untagged")
os.makedirs(tmp, exist_ok=True)
subprocess.run(["soffice", "--headless", "--convert-to", untagged, "--outdir", tmp,
                os.path.join(OUT, "synthetic-report.docx")], check=True, capture_output=True)
os.replace(os.path.join(tmp, "synthetic-report.pdf"), os.path.join(OUT, "synthetic-report-untagged.pdf"))
os.rmdir(tmp)
print("fixtures written to", OUT)
