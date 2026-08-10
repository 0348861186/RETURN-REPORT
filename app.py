import io
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

st.set_page_config(page_title="CCR Generator", layout="centered")

st.title("📊 CCR Report Generator / Tạo Báo Cáo Khiếu Nại")
st.write("Tải file Excel để tự động tạo báo cáo PPTX song ngữ.")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])


def process_dataframe(df):
    """Làm sạch tên cột từ file Excel có xuống dòng/tiếng Việt"""
    # Gộp/Chuẩn hóa tên cột bằng cách bỏ các ký tự xuống dòng và khoảng trắng thừa
    df.columns = [str(col).replace('\n', ' ').strip() for col in df.columns]

    # Hàm tìm tên cột chính xác dù tiêu đề chứa cả Tiếng Anh + Tiếng Việt
    def find_col(keywords):
        for col in df.columns:
            if any(kw.lower() in col.lower() for kw in keywords):
                return col
        return None

    col_map = {
        'code': find_col(['Complaint Code', 'Mã khiếu nại']),
        'date': find_col(['Date', 'Ngày']),
        'facility': find_col(['Facility', 'Bộ phận']),
        'customer': find_col(['Customer', 'Khách hàng']),
        'defect': find_col(['Defect type', 'Loại lỗi']),
        'remarks': find_col(['Remarks']),
        'res_rate': find_col(['Resolution rate', 'Tỉ lệ hoàn thành']),
        'tat': find_col(['AVR turnaround time', 'Thời gian điều tra'])
    }
    return df, col_map


def generate_charts(df, col_map):
    """Tạo biểu đồ Trend và Defect Category từ dữ liệu Excel"""
    charts = {}

    # 1. Biểu đồ Trend
    if col_map['date'] and col_map['date'] in df.columns:
        fig, ax = plt.subplots(figsize=(5, 3))
        df_trend = df.copy()
        df_trend['parsed_date'] = pd.to_datetime(df_trend[col_map['date']], errors='coerce')
        weekly_counts = df_trend.groupby(df_trend['parsed_date'].dt.isocalendar().week).size()

        if not weekly_counts.empty:
            ax.plot([f"W{w}" for w in weekly_counts.index], weekly_counts.values, marker='o', color='#1e3d59', linewidth=2)
        ax.set_title("Weekly Complaint Trend", fontsize=10)
        plt.tight_layout()

        trend_img = io.BytesIO()
        plt.savefig(trend_img, format='png', dpi=150)
        trend_img.seek(0)
        charts['trend'] = trend_img
        plt.close()

    # 2. Biểu đồ Defect Category
    if col_map['defect'] and col_map['defect'] in df.columns:
        fig, ax = plt.subplots(figsize=(5, 3))
        defect_counts = df[col_map['defect']].astype(str).value_counts()

        ax.bar(defect_counts.index, defect_counts.values, color='#ff6e40')
        ax.set_title("Defect Category Distribution", fontsize=10)
        plt.xticks(rotation=15, ha='right', fontsize=8)
        plt.tight_layout()

        defect_img = io.BytesIO()
        plt.savefig(defect_img, format='png', dpi=150)
        defect_img.seek(0)
        charts['defect'] = defect_img
        plt.close()

    return charts


def create_pptx(df, col_map, charts):
    """Tạo Presentation 5 slide song ngữ"""
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: Title ---
    slide1 = prs.slides.add_slide(blank_layout)
    txBox = slide1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(2))
    tf = txBox.text_frame
    p1 = tf.paragraphs[0]
    p1.text = "WEEKLY CUSTOMER COMPLAINT REPORT (CCR)"
    p1.font.bold = True
    p1.font.size = Pt(32)
    p1.font.color.rgb = RGBColor(30, 61, 89)

    p2 = tf.add_paragraph()
    p2.text = "BÁO CÁO KHIẾU NẠI KHÁCH HÀNG HÀNG TUẦN"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(100, 100, 100)

    # --- SLIDE 2: Overview / Key Metrics ---
    slide2 = prs.slides.add_slide(blank_layout)
    tb = slide2.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    p = tb.text_frame.paragraphs[0]
    p.text = "Overview / Tổng quan"
    p.font.size = Pt(24)
    p.font.bold = True

    total_cases = len(df)
    
    # Lấy giá trị KPI từ dòng có dữ liệu đầu tiên
    res_rate_col = col_map['res_rate']
    res_rate = "N/A"
    if res_rate_col and res_rate_col in df.columns:
        valid_res = df[res_rate_col].dropna()
        if not valid_res.empty:
            res_val = valid_res.iloc[0]
            res_rate = f"{float(res_val)*100:.0f}%" if isinstance(res_val, (int, float)) else str(res_val)

    tat_col = col_map['tat']
    tat = "N/A"
    if tat_col and tat_col in df.columns:
        valid_tat = df[tat_col].dropna()
        if not valid_tat.empty:
            tat = f"{valid_tat.iloc[0]} Days"

    metrics = [
        ("TOTAL COMPLAINTS\nTỔNG SỐ KHIẾU NẠI", str(total_cases)),
        ("RESOLUTION RATE\nTỶ LỆ HOÀN THÀNH", str(res_rate)),
        ("AVG TURNAROUND TIME\nTHỜI GIAN ĐIỀU TRA TB", str(tat))
    ]

    for i, (title, val) in enumerate(metrics):
        shape = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1 + i * 3.8), Inches(2), Inches(3.5), Inches(2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(240, 244, 248)
        tf = shape.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(12)
        p.font.color.rgb = RGBColor(100, 100, 100)
        p_val = tf.add_paragraph()
        p_val.text = val
        p_val.font.size = Pt(28)
        p_val.font.bold = True
        p_val.font.color.rgb = RGBColor(30, 61, 89)

    # --- SLIDE 3: Trends & Defect Category ---
    slide3 = prs.slides.add_slide(blank_layout)
    tb = slide3.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    tb.text_frame.paragraphs[0].text = "Trend & Defect Category / Xu hướng & Phân loại lỗi"

    if 'trend' in charts:
        slide3.shapes.add_picture(charts['trend'], Inches(1), Inches(1.8), width=Inches(5.5))
    if 'defect' in charts:
        slide3.shapes.add_picture(charts['defect'], Inches(6.8), Inches(1.8), width=Inches(5.5))

    # --- SLIDE 4: Pareto Analysis ---
    slide4 = prs.slides.add_slide(blank_layout)
    tb = slide4.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    tb.text_frame.paragraphs[0].text = "Pareto Analysis (80/20) / Phân tích Pareto"

    if 'defect' in charts:
        slide4.shapes.add_picture(charts['defect'], Inches(1), Inches(1.8), width=Inches(5.5))

    txBox = slide4.shapes.add_textbox(Inches(6.8), Inches(2), Inches(5.5), Inches(4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Key Takeaways / Điểm chính:"
    p.font.bold = True
    p.font.size = Pt(16)

    p2 = tf.add_paragraph()
    p2.text = "• Shortage (Thiếu số lượng) accounts for the highest proportion.\n  Lỗi thiếu số lượng chiếm tỷ trọng cao nhất."
    p2.font.size = Pt(14)

    # --- SLIDE 5: Pending Cases Analysis ---
    slide5 = prs.slides.add_slide(blank_layout)
    tb = slide5.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    tb.text_frame.paragraphs[0].text = "Pending Cases Analysis / Phân tích các trường hợp đang xử lý"

    remarks_col = col_map['remarks']
    if remarks_col and remarks_col in df.columns:
        pending_df = df[df[remarks_col].astype(str).str.contains('Pending', case=False, na=False)]
    else:
        pending_df = pd.DataFrame()

    rows = len(pending_df) + 1
    cols = 5
    table_shape = slide5.shapes.add_table(rows, cols, Inches(0.8), Inches(1.8), Inches(11.7), Inches(0.5 * rows))
    table = table_shape.table

    headers = ["Complaint Code\nMã khiếu nại", "Customer\nKhách hàng", "Defect Type\nLoại lỗi", "Facility\nBộ phận", "Status\nTrạng thái"]
    for i, h in enumerate(headers):
        table.cell(0, i).text = h

    for row_idx, (_, row) in enumerate(pending_df.iterrows(), start=1):
        table.cell(row_idx, 0).text = str(row.get(col_map['code'], ''))
        table.cell(row_idx, 1).text = str(row.get(col_map['customer'], ''))
        table.cell(row_idx, 2).text = str(row.get(col_map['defect'], ''))
        table.cell(row_idx, 3).text = str(row.get(col_map['facility'], ''))
        table.cell(row_idx, 4).text = str(row.get(col_map['remarks'], ''))

    pptx_out = io.BytesIO()
    prs.save(pptx_out)
    pptx_out.seek(0)
    return pptx_out


# Main Streamlit App Logic
if uploaded_file:
    try:
        raw_df = pd.read_excel(uploaded_file)
        df, col_map = process_dataframe(raw_df)

        st.success("File Excel loaded successfully / Đã tải dữ liệu thành công!")
        st.dataframe(df.head(3))

        charts = generate_charts(df, col_map)
        pptx_data = create_pptx(df, col_map, charts)

        st.subheader("📥 Download Report / Tải Báo Cáo")
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=pptx_data,
            file_name="Weekly_Customer_Complaint_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        st.error(f"Lỗi xử lý file Excel: {str(e)}")
