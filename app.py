from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# 1. Khởi tạo Presentation (Màn hình rộng 16:9)
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6] # Blank slide layout

# Bảng màu chủ đạo (Color Palette)
COLOR_BG = RGBColor(245, 247, 250)      # Nền sáng
COLOR_BLUE = RGBColor(37, 99, 235)      # Xanh lam
COLOR_GREEN = RGBColor(16, 185, 129)    # Xanh lá
COLOR_ORANGE = RGBColor(245, 158, 11)   # Cam
COLOR_RED = RGBColor(220, 38, 38)       # Đỏ
COLOR_DARK = RGBColor(30, 41, 59)       # Chữ đen xám
COLOR_WHITE = RGBColor(255, 255, 255)  # Trắng


# -----------------------------------------------------------------------------
# SLIDE 3: OVERVIEW & KEY METRICS
# -----------------------------------------------------------------------------
slide_3 = prs.slides.add_slide(blank_layout)

# Banner Tiêu Đề Top
banner = slide_3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.4), Inches(12.333), Inches(0.8))
banner.fill.solid()
banner.fill.fore_color.rgb = COLOR_BLUE
banner.line.color.rgb = COLOR_BLUE
tf_b = banner.text_frame
tf_b.text = "OVERVIEW & KEY METRICS\nTỔNG QUAN & CHỈ SỐ KPI HÀNG TUẦN"
tf_b.paragraphs[0].font.size = Pt(14)
tf_b.paragraphs[0].font.bold = True
tf_b.paragraphs[0].font.color.rgb = COLOR_WHITE

# Thẻ Metric 1 (Total Complaints)
card1 = slide_3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.4), Inches(3.9), Inches(1.8))
card1.fill.solid()
card1.fill.fore_color.rgb = COLOR_WHITE
card1.line.color.rgb = COLOR_BLUE
tf1 = card1.text_frame
tf1.text = "3 Cases\nTOTAL COMPLAINTS RECEIVED\nTổng số khiếu nại phát sinh"
tf1.paragraphs[0].font.size = Pt(28)
tf1.paragraphs[0].font.bold = True
tf1.paragraphs[0].font.color.rgb = COLOR_BLUE

# Bảng Danh Sách Complaint
table_shape = slide_3.shapes.add_table(4, 4, Inches(0.5), Inches(3.4), Inches(7.0), Inches(2.5))
table = table_shape.table
headers = ["Complaint Code", "Facility", "Customer", "Status"]
for i, name in enumerate(headers):
    table.cell(0, i).text = name


# -----------------------------------------------------------------------------
# SLIDE 4: TREND & DEFECT CATEGORY ANALYSIS (Donut Chart)
# -----------------------------------------------------------------------------
slide_4 = prs.slides.add_slide(blank_layout)

# Banner Tiêu Đề Đỏ
banner4 = slide_4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.4), Inches(12.333), Inches(0.8))
banner4.fill.solid()
banner4.fill.fore_color.rgb = COLOR_RED
banner4.line.color.rgb = COLOR_RED
tf_b4 = banner4.text_frame
tf_b4.text = "PARETO 80/20 ANALYSIS\nNGUYÊN TẮC PARETO - NỔI BẬT LỖI TRỌNG YẾU"
tf_b4.paragraphs[0].font.bold = True
tf_b4.paragraphs[0].font.color.rgb = COLOR_WHITE

# Biểu đồ Donut (Defect Breakdown)
chart_data = CategoryChartData()
chart_data.categories = ['Shortage', 'Others']
chart_data.add_series('Defects', (66.7, 33.3))

x, y, cx, cy = Inches(6.8), Inches(1.5), Inches(5.5), Inches(3.5)
chart = slide_4.shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, chart_data
).chart
chart.has_legend = True

# Lưu file PPTX
prs.save("Weekly_CCR_Report.pptx")
print("Đã xuất file PowerPoint thành công!")
