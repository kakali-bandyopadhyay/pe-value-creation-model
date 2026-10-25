from __future__ import annotations

"""Render compact PNG previews from recalculated workbook cells."""

from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "model" / "PE_Value_Creation_Model_synthetic.xlsx"
OUT = ROOT / "docs"
NAVY = "#17365D"
BLUE = "#D9EAF7"
PALE = "#EAF2F8"
GREEN = "#E2F0D9"
AMBER = "#FFF2CC"
RED = "#FCE4D6"
INK = "#202124"
GREY = "#667085"
WHITE = "#FFFFFF"
GRID = "#C7D3DE"


def fonts():
    path = Path(r"C:\Windows\Fonts\arial.ttf")
    bold_path = Path(r"C:\Windows\Fonts\arialbd.ttf")
    try:
        return (ImageFont.truetype(str(path), 17), ImageFont.truetype(str(bold_path), 18),
                ImageFont.truetype(str(path), 14), ImageFont.truetype(str(path), 12))
    except OSError:
        return (ImageFont.load_default(),) * 4


def fmt(value, number_format=None):
    if value is None:
        return ""
    if isinstance(value, float):
        if number_format and "%" in number_format:
            return f"{value:.1%}"
        if number_format and "x" in number_format.lower():
            return f"{value:.2f}x"
        if number_format and "$" in number_format:
            return f"${value:,.1f}"
        return f"{value:,.1f}"
    return str(value)


def rounded(draw, box, fill, radius=10, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def render_dashboard(wb):
    ws = wb["Dashboard"]
    image = Image.new("RGB", (1700, 1100), "#F4F7FA")
    draw = ImageDraw.Draw(image)
    body, bold, small, tiny = fonts()
    draw.rectangle((0, 0, 1700, 112), fill=NAVY)
    draw.text((54, 24), "NimbleBridge Services (synthetic)", font=bold, fill=WHITE)
    draw.text((54, 62), "Value creation and sponsor returns | synthetic workbook preview", font=small, fill="#D8E6F2")
    scenario = ws["D5"].value or "Base"
    draw.text((1380, 38), f"Selected case: {scenario}", font=body, fill=WHITE)
    cards = [("Entry EV", ws["D9"].value, '$#,##0.0'), ("Sponsor equity", ws["D10"].value, '$#,##0.0'),
             ("Exit EBITDA", ws["D11"].value, '$#,##0.0'), ("MOIC", ws["D14"].value, '0.0x'), ("IRR", ws["D15"].value, '0.0%')]
    x0, card_w, gap = 54, 300, 25
    for i, (label, val, f) in enumerate(cards):
        x = x0 + i * (card_w + gap)
        rounded(draw, (x, 138, x + card_w, 245), WHITE, outline=GRID)
        draw.text((x + 20, 154), label, font=small, fill=GREY)
        draw.text((x + 20, 190), fmt(val, f), font=bold, fill=NAVY)
    # Trend chart based on the Dashboard tab's historical and forecast rows.
    rounded(draw, (54, 273, 855, 655), WHITE, outline=GRID)
    draw.text((78, 292), "Revenue and EBITDA (USD millions)", font=bold, fill=INK)
    years = [ws.cell(35, col).value for col in range(4, 12)]
    revenue = [float(ws.cell(36, col).value or 0) for col in range(4, 12)]
    ebitda = [float(ws.cell(37, col).value or 0) for col in range(4, 12)]
    left, top, right, bottom = 105, 350, 820, 590
    draw.line((left, bottom, right, bottom), fill=GRID, width=2)
    max_value = max(revenue + ebitda) * 1.1
    def xy(i, value):
        return (left + i * (right - left) / (len(years)-1), bottom - value / max_value * (bottom-top))
    for series, color, name in ((revenue, "#1F77B4", "Revenue"), (ebitda, "#E07A2D", "EBITDA")):
        points = [xy(i, val) for i, val in enumerate(series)]
        draw.line(points, fill=color, width=5, joint="curve")
        for px, py in points:
            draw.ellipse((px-5, py-5, px+5, py+5), fill=color)
    for i, year in enumerate(years):
        px, _ = xy(i, 0)
        draw.text((px-22, bottom+10), str(year), font=tiny, fill=GREY)
    draw.line((112, 325, 145, 325), fill="#1F77B4", width=5); draw.text((154, 316), "Revenue", font=tiny, fill=INK)
    draw.line((255, 325, 288, 325), fill="#E07A2D", width=5); draw.text((297, 316), "EBITDA", font=tiny, fill=INK)
    # EBITDA bridge.
    rounded(draw, (885, 273, 1645, 655), WHITE, outline=GRID)
    draw.text((910, 292), "FY2025 baseline to FY2030 EBITDA", font=bold, fill=INK)
    bridge_rows = [(19, "FY2025 baseline"), (20, "Core revenue and margin change"),
                   (21, "Pricing and mix"), (22, "Commercial efficiency"),
                   (23, "AI finance automation"), (24, "FY2030 exit EBITDA")]
    y = 346
    for row, label in bridge_rows:
        val = ws[f"D{row}"].value
        draw.text((914, y), label, font=small if row not in (19,24) else body, fill=INK)
        draw.text((1515, y), fmt(val, '$#,##0.0'), font=body, fill=NAVY)
        y += 47
    # Value plan summary and exit value bridge.
    rounded(draw, (54, 680, 820, 1020), WHITE, outline=GRID)
    draw.text((78, 700), "Quarterly plan tracking", font=bold, fill=INK)
    for i, row in enumerate((41, 42, 43, 44, 45)):
        draw.text((80, 753 + i*45), str(ws[f"C{row}"].value), font=small, fill=INK)
        draw.text((710, 753 + i*45), fmt(ws[f"D{row}"].value, '#,##0'), font=body, fill=NAVY)
    rounded(draw, (885, 680, 1645, 1020), WHITE, outline=GRID)
    draw.text((910, 700), "Enterprise value bridge", font=bold, fill=INK)
    for i, row in enumerate((28,29,30,31)):
        draw.text((914, 765 + i*52), str(ws[f"C{row}"].value), font=small, fill=INK)
        draw.text((1515, 765 + i*52), fmt(ws[f"D{row}"].value, '$#,##0.0'), font=body, fill=NAVY)
    draw.text((54, 1050), "Rendered preview from the recalculated workbook. All displayed figures are synthetic and illustrative.", font=tiny, fill=GREY)
    image.save(OUT / "dashboard_preview.png")


def render_tracker(wb):
    ws = wb["VCP Tracker"]
    image = Image.new("RGB", (2200, 1060), "#F4F7FA")
    draw = ImageDraw.Draw(image)
    body, bold, small, tiny = fonts()
    draw.rectangle((0, 0, 2200, 115), fill=NAVY)
    draw.text((48, 24), "Quarterly value creation plan tracker", font=bold, fill=WHITE)
    draw.text((48, 63), "Synthetic plan and synthetic actuals | accountable role, KPI, variance, status and next action", font=small, fill="#D8E6F2")
    columns = [(3, "Initiative", 240), (4,"Quarter",120), (5,"Owner role",310), (6,"KPI",230),
               (7,"Plan",110), (8,"Synthetic actual",150), (9,"Variance",130), (10,"RAG",100),
               (11,"Quality KPI",120), (12,"AI gate",110), (13,"Next action",380)]
    start_x, y0, row_h = 38, 162, 67
    x_positions=[]
    x=start_x
    for col,label,width in columns:
        x_positions.append((col,x,width,label))
        rounded(draw,(x,y0,x+width,y0+56),NAVY,radius=4)
        draw.text((x+10,y0+17),label,font=small,fill=WHITE)
        x += width+4
    for row in range(8,20):
        y=y0+58+(row-8)*row_h
        if y+row_h>1015:break
        fill=WHITE if row%2 else "#F8FBFD"
        draw.rectangle((start_x,y,x-4,y+row_h-4),fill=fill,outline=GRID,width=1)
        for col,left,width,label in x_positions:
            val=ws.cell(row,col).value
            text=fmt(val,ws.cell(row,col).number_format)
            color=INK
            if col==10:
                fill_status={"Green":GREEN,"Amber":AMBER,"Red":RED}.get(text,fill)
                draw.rounded_rectangle((left+8,y+15,left+width-8,y+row_h-16),radius=8,fill=fill_status)
                color=INK
            if col==12:
                fill_gate={"Expand":GREEN,"Fix":AMBER,"Pause":RED}.get(text,fill)
                draw.rounded_rectangle((left+7,y+15,left+width-7,y+row_h-16),radius=8,fill=fill_gate)
                color=INK
            words=text.split()
            max_chars=max(10,int(width/9))
            lines=[];current=""
            for word in words:
                if len(current)+len(word)+1>max_chars and current:
                    lines.append(current);current=word
                else: current=(current+" "+word).strip()
            if current:lines.append(current)
            for j,line in enumerate(lines[:2]):
                draw.text((left+9,y+11+j*20),line,font=tiny if len(text)>25 else small,fill=color)
    draw.text((38, 1018), "Preview from workbook rows 8–19. The full tracker contains five years of quarterly synthetic records.", font=tiny, fill=GREY)
    draw.text((38, 1038), "AI gate: Pause takes precedence if adoption is below threshold; otherwise low quality → Fix; both pass → Expand. Governance signal only; forecast is unchanged.", font=tiny, fill=GREY)
    image.save(OUT / "vcp_tracker_preview.png")


def main():
    if not WORKBOOK.exists():
        raise FileNotFoundError("Build and recalculate the workbook before rendering previews.")
    wb = load_workbook(WORKBOOK, data_only=True, read_only=True)
    render_dashboard(wb)
    render_tracker(wb)
    print(OUT / "dashboard_preview.png")
    print(OUT / "vcp_tracker_preview.png")


if __name__ == "__main__":
    main()
