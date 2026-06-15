from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "output" / "ridgeline_freight_accrual_audit_pack.xlsx"
OUT = ROOT / "assets" / "accrual_breakdown.png"


def main():
    df = pd.read_excel(WORKBOOK, sheet_name="Carrier Accrual Detail")
    df = df.sort_values("estimated_total", ascending=True)

    width, height = 940, 620
    image = Image.new("RGB", (width, height), "#ffffff")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()

    colors = {
        "Peak Logistics": "#be5b38",
        "Heartland Freight": "#0f6b5c",
        "Coastal Express": "#315d89",
    }

    draw.text((48, 38), "April 2026 Freight Accrual by Carrier", fill="#17202a", font=font)
    total = float(df["estimated_total"].sum())
    draw.text((48, 72), f"Total accrual: ${total:,.2f}", fill="#52616f", font=font)

    max_value = float(df["estimated_total"].max())
    y = 120
    bar_x = 230
    bar_width = 610
    row_height = 105

    for _, row in df.iterrows():
        carrier = str(row["carrier"])
        value = float(row["estimated_total"])
        filled_width = int(bar_width * value / max_value)

        draw.text((48, y + 18), carrier, fill="#17202a", font=font)
        draw.rounded_rectangle((bar_x, y, bar_x + bar_width, y + 38), radius=8, fill="#e6eaee")
        draw.rounded_rectangle(
            (bar_x, y, bar_x + filled_width, y + 38),
            radius=8,
            fill=colors.get(carrier, "#0f6b5c"),
        )
        draw.text((bar_x + bar_width - 105, y + 52), f"${value:,.0f}", fill="#17202a", font=font)
        draw.text((bar_x, y + 52), f"Risk: {row['confidence_risk_rating']}", fill="#52616f", font=font)
        y += row_height

    draw.line((48, 520, 892, 520), fill="#d7dde3", width=1)
    draw.text(
        (48, 545),
        "Generated from ridgeline_freight_accrual_audit_pack.xlsx",
        fill="#52616f",
        font=font,
    )

    OUT.parent.mkdir(exist_ok=True)
    image.save(OUT)


if __name__ == "__main__":
    main()
