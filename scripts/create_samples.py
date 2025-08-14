from __future__ import annotations
from pathlib import Path
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image, ImageDraw, ImageFont


BASE_DIR = Path(r"data")  # change to your folder
BASE_DIR.mkdir(parents=True, exist_ok=True)


def create_excel_sample(path: Path) -> None:
    """Create a small Excel file with fake weekly data."""
    data = {
        "Day": ["Mon", "Tue", "Wed", "Thu", "Fri"],
        "Sales": [1200, 1350, 980, 1500, 1420],
        "Leads": [12, 15, 9, 18, 14],
    }
    df = pd.DataFrame(data)
    df.to_excel(path, index=False)
    print(f"✅ Excel sample saved at {path}")


def create_pdf_sample(path: Path) -> None:
    """Create a sample PDF file."""
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(72, 720, "Weekly Summary Report")
    c.setFont("Helvetica", 12)
    c.drawString(72, 690, "This is an auto-generated summary for testing email automation.")
    c.drawString(72, 670, "Sales performance improved over last week by 8%.")
    c.showPage()
    c.save()
    print(f"✅ PDF sample saved at {path}")


def create_banner_image(path: Path) -> None:
    """Create a simple banner image."""
    img = Image.new("RGB", (600, 150), color="#004aad")
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 36)
    except IOError:
        font = ImageFont.load_default()
    draw.text((50, 50), "Weekly Report", font=font, fill="white")
    img.save(path)
    print(f"✅ Banner image saved at {path}")


if __name__ == "__main__":
    create_excel_sample(Path(r"data/reports/weekly.xlsx"))
    create_pdf_sample(Path(r"data/reports/summary.pdf"))
    create_banner_image(Path(r"data/assets/banner.png"))
