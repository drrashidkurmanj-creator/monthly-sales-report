from fpdf import FPDF
from openpyxl import load_workbook
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

# --- Settings: change only these two lines ---
input_file = "oktober_daten.xlsx"
report_title = "Oktober 2026"
# ---------------------------------------------

wb = load_workbook(input_file)
ws = wb.active

by_category = {}
by_article = {}
total = 0

for row in ws.iter_rows(min_row=2, values_only=True):
    category = row[1]
    article = row[2]
    line = row[3] * row[4]

    total = total + line
    by_category[category] = by_category.get(category, 0) + line
    by_article[article] = by_article.get(article, 0) + line

top = sorted(by_article, key=by_article.get, reverse=True)

# chart
names = sorted(by_category)
values = []
for n in names:
    values.append(by_category[n])

plt.figure(figsize=(6, 4))
plt.bar(names, values)
plt.title(f"Umsatz nach Kategorie - {report_title}")
plt.ylabel("EUR")
plt.tight_layout()
plt.savefig("umsatz_kategorie.png", dpi=150)
plt.close()

# PDF
pdf = FPDF()
pdf.add_page()

pdf.set_font("Helvetica", "B", 20)
pdf.cell(0, 12, f"Monatsbericht {report_title}", new_x="LMARGIN", new_y="NEXT")

pdf.set_font("Helvetica", size=12)
pdf.cell(0, 8, f"Gesamtumsatz: {total:.2f} EUR", new_x="LMARGIN", new_y="NEXT")
pdf.ln(4)

pdf.image("umsatz_kategorie.png", w=150)
pdf.ln(6)


def table_row(a, b, c, bold=False):
    style = "B" if bold else ""
    pdf.set_font("Helvetica", style=style, size=11)
    pdf.cell(70, 8, a, border=1)
    pdf.cell(50, 8, b, border=1, align="R")
    pdf.cell(40, 8, c, border=1, align="R", new_x="LMARGIN", new_y="NEXT")


table_row("Kategorie", "Umsatz (EUR)", "Anteil", bold=True)

for category in sorted(by_category):
    value = by_category[category]
    share = value / total * 100
    table_row(category, f"{value:.2f}", f"{share:.1f} %")

pdf.ln(6)
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 8, "Top 3 Artikel", new_x="LMARGIN", new_y="NEXT")

table_row("Artikel", "Umsatz (EUR)", "Anteil", bold=True)

for name in top[:3]:
    value = by_article[name]
    share = value / total * 100
    table_row(name, f"{value:.2f}", f"{share:.1f} %")

pdf_name = f"Monatsbericht_{report_title.replace(' ', '_')}.pdf"
pdf.output(pdf_name)
print("PDF created:", pdf_name)