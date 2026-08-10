import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io

# Cấu hình trang Streamlit
st.set_page_config(page_title="CCR Generator", layout="centered")

st.title("📊 CCR Report Generator / Tạo Báo Cáo Khiếu Nại")
st.write("Tải file Excel để tự động tạo báo cáo PPTX song ngữ.")

# 1. Tải file Excel
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

def clean_df_headers(df):
    """
    Xóa xuống dòng (\n) và khoảng trắng thừa trong tiêu đề cột.
    Ví dụ: 'Date\nNgày' -> 'Date'
    """
    new_columns = {}
    for col in df.columns:
        clean_name = str(col).split('\n')[0].strip()
        new_columns[col] = clean_name
    df = df.rename(columns=new_columns)
    return df

def generate_charts(df):
    """Tạo biểu đồ Trend và Pareto từ dữ liệu Excel"""
    charts = {}
    
    # --- Biểu đồ Trend (Slide 3) ---
    fig, ax = plt.subplots(figsize=(5, 3))
    
    if 'Date' in df.columns:
        clean_df = df.dropna(subset=['Date']).copy()
        clean_df['Parsed_Date'] = pd.to_datetime(clean_df['Date'], errors='coerce')
        clean_df = clean_df.dropna(subset=['Parsed_Date'])
        
        if not clean_df.empty:
            weekly_counts = clean_df.groupby(clean_df['Parsed_Date'].dt.isocalendar().week).size()
            ax.plot([f"W{w}" for w in weekly_counts.index], weekly_counts.values, marker='o', color='#1e3d59', linewidth=2)
            ax.set_title("Weekly Complaint Trend", fontsize=10, fontweight='bold')
        else:
            ax.text(0.5, 0.5, "Không có dữ liệu ngày hợp lệ", ha='center', va='center')
    else:
        ax.text(0.5, 0.5, "Thiếu cột 'Date'", ha='center', va='center')
        
    plt.tight_layout()
    trend_img = io.BytesIO()
    plt.savefig(trend_img, format='png', dpi=150)
    trend_img.seek(0)
    charts['trend'] = trend_img
    plt.close()

    # --- Biểu đồ Defect Category Pareto (Slide 3 & 4) ---
    fig, ax = plt.subplots(figsize=(5, 3))
    
    if 'Defect type' in df.columns:
        defect_counts = df['Defect type'].value_counts().head(7)
        ax.bar(defect_counts.index, defect_counts.values, color='#ff6e40')
        ax.set_title("Defect Category Distribution", fontsize=10, fontweight='bold')
        plt.xticks(rotation=20, ha='right', fontsize=8)
    else:
        ax.text(0.5, 0.5, "Thiếu cột 'Defect type'", ha='center', va='center')
        
    plt.tight_layout()
    defect_img = io.BytesIO()
    plt.savefig(defect_img, format='png', dpi=150)
    defect_img.seek(0)
    charts['defect'] = defect_img
    plt.close()

    return charts

def create_pptx(df, charts):
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
    
    # Lấy thông số KPI
    total_cases = len(df)
    
    res_val = df['Resolution rate'].dropna().iloc[0] if 'Resolution rate' in df.columns and not df['Resolution rate'].dropna().empty else "N/A"
    if isinstance(res_val, float):
        res_rate = f"{res_val * 100:.0f}%" if res_val <= 1 else f"{res_val:.0f}%"
    else:
        res_rate = str(res_val)

    tat_val = df['AVR turnaround time'].dropna().iloc[0] if 'AVR turnaround time' in df.columns and not df['AVR turnaround time'].dropna().empty else "N/A"
    tat = f"{tat_val} Days" if tat_val != "N/A" else "N/A"

    metrics = [
        ("TOTAL COMPLAINTS\nTỔNG SỐ KHIẾU NẠI", str(total_cases)),
        ("RESOLUTION RATE\nTỶ LỆ HOÀN THÀNH", res_rate),
        ("AVG TURNAROUND TIME\nTHỜI GIAN ĐIỀU TRA TB", tat)
    ]
    
    for i, (title, val) in enumerate(metrics):
        shape = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1 + i*3.8), Inches(2.2), Inches(3.5), Inches(2.2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(240, 244, 248)
        tf = shape.text_frame
        
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(12)
        p.font.color.rgb = RGBColor(100, 100, 100)
        
        p_val = tf.add_paragraph()
        p_val.text = val
        p_val.alignment = PP_ALIGN.CENTER
        p_val.font.size = Pt(28)
        p_val.font.bold = True
        p_val.font.color.rgb = RGBColor(30, 61, 89)

    # --- SLIDE 3: Trends & Defect Category ---
    slide3 = prs.slides.add_slide(blank_layout)
    tb = slide3.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    p = tb.text_frame.paragraphs[0]
    p.text = "Trend & Defect Category / Xu hướng & Phân loại lỗi"
    p.font.size = Pt(24)
    p.font.bold = True
    
    slide3.shapes.add_picture(charts['trend'], Inches(0.8), Inches(1.8), width=Inches(5.6))
    slide3.shapes.add_picture(charts['defect'], Inches(6.8), Inches(1.8), width=Inches(5.6))

    # --- SLIDE 4: Pareto Analysis (80/20) ---
    slide4 = prs.slides.add_slide(blank_layout)
    tb = slide4.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    p = tb.text_frame.paragraphs[0]
    p.text = "Pareto Analysis (80/20) / Phân tích Pareto"
    p.font.size = Pt(24)
    p.font.bold = True
    
    slide4.shapes.add_picture(charts['defect'], Inches(0.8), Inches(1.8), width=Inches(5.6))
    
    txBox = slide4.shapes.add_textbox(Inches(6.8), Inches(2), Inches(5.5), Inches(4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Key Takeaways / Điểm chính:"
    p.font.bold = True
    p.font.size = Pt(16)
    
    p2 = tf.add_paragraph()
    p2.text = "• Main defect categories account for majority of total complaints.\n  Các loại lỗi chính chiếm phần lớn tổng số trường hợp khiếu nại."
    p2.font.size = Pt(14)

    # --- SLIDE 5: Pending Cases Analysis ---
    slide5 = prs.slides.add_slide(blank_layout)
    tb = slide5.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    p = tb.text_frame.paragraphs[0]
    p.text = "Pending Cases Analysis / Phân tích các trường hợp đang xử lý"
    p.font.size = Pt(24)
    p.font.bold = True
    
    # Lọc danh sách Pending
    if 'Remarks' in df.columns:
        pending_df = df[df['Remarks'].astype(str).str.contains('Pending', case=False, na=False)].head(8)
    else:
        pending_df = pd.DataFrame()
    
    rows = max(len(pending_df) + 1, 2)
    cols = 5
    table_shape = slide5.shapes.add_table(rows, cols, Inches(0.8), Inches(1.8), Inches(11.7), Inches(0.5 * rows))
    table = table_shape.table
    
    headers = ["Complaint Code\nMã khiếu nại", "Customer\nKhách hàng", "Defect Type\nLoại lỗi", "Facility\nBộ phận", "Status\nTrạng thái"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(10)
            paragraph.font.bold = True
        
    if not pending_df.empty:
        for row_idx, (_, row) in enumerate(pending_df.iterrows(), start=1):
            vals = [
                str(row.get('Complaint Code', '')),
                str(row.get('Customer', '')),
                str(row.get('Defect type', '')),
                str(row.get('Facility', '')),
                str(row.get('Remarks', ''))
            ]
            for col_idx, val in enumerate(vals):
                cell = table.cell(row_idx, col_idx)
                cell.text = val
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(9)

    # Lưu PPTX vào bộ nhớ BytesIO
    pptx_out = io.BytesIO()
    prs.save(pptx_out)
    pptx_out.seek(0)
    return pptx_out

# Xử lý chính trong ứng dụng Streamlit
if uploaded_file:
    try:
        # Đọc dữ liệu từ Dòng 3 (header=2) để bỏ qua tiêu đề trang trí "CUSTOMER COMPLAINT STATISTICS"
        df = pd.read_excel(uploaded_file, header=2)
        
        # Làm sạch tên cột (loại bỏ phần dịch tiếng Việt bên dưới dấu \n)
        df = clean_df_headers(df)
        
        # Bỏ các dòng hoàn toàn trống
        df = df.dropna(how='all')
        
        st.success("Đã tải và xử lý file Excel thành công!")
        
        # Tạo biểu đồ và PPTX
        charts = generate_charts(df)
        pptx_data = create_pptx(df, charts)

        # Nút Download PPTX
        st.subheader("📥 Download Report / Tải Báo Cáo")
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=pptx_data,
            file_name="Weekly_Customer_Complaint_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        st.info("💡 Mẹo: Bạn có thể mở file PPTX đã tải và dùng tính năng 'Save As' trong Microsoft PowerPoint nếu muốn xuất bản PDF.")
        
    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi đọc/xử lý file Excel: {e}")
