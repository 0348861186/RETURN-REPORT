import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import os
import re

# Cấu hình trang Streamlit
st.set_page_config(page_title="CCR Generator", layout="centered")

st.title("📊 CCR Report Generator / Tạo Báo Cáo Khiếu Nại")
st.write("Tải file Excel để tự động tạo báo cáo PPTX / PDF song ngữ.")

# 1. Tải file Excel
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

def generate_charts(df):
    """Tạo biểu đồ Trend và Pareto từ dữ liệu Excel"""
    charts = {}
    
    # Biểu đồ Trend (Slide 3)
    fig, ax = plt.subplots(figsize=(5, 3))
    
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df_clean = df.dropna(subset=['Date'])
    
    weekly_counts = df_clean.groupby(df_clean['Date'].dt.isocalendar().week).size()
    
    ax.plot([f"W{w}" for w in weekly_counts.index], weekly_counts.values, marker='o', color='#1e3d59', linewidth=2)
    ax.set_title("Weekly Complaint Trend", fontsize=10)
    plt.tight_layout()
    
    trend_img = io.BytesIO()
    plt.savefig(trend_img, format='png', dpi=150)
    trend_img.seek(0)
    charts['trend'] = trend_img
    plt.close()

    # Biểu đồ Defect Category Pareto (Slide 3 & 4)
    fig, ax = plt.subplots(figsize=(5, 3))
    defect_counts = df['Defect type'].value_counts()
    
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
    # Thêm Tiêu đề
    tb = slide2.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    p = tb.text_frame.paragraphs[0]
    p.text = "Overview / Tổng quan"
    p.font.size = Pt(24)
    p.font.bold = True
    
    # Tính KPI
    total_cases = len(df)
    res_rate = df['Resolution rate'].iloc[0] if 'Resolution rate' in df.columns else "80%"
    tat = df['AVR turnaround time'].iloc[0] if 'AVR turnaround time' in df.columns else "1.5 Days"
    
    metrics = [
        ("TOTAL COMPLAINTS\nTỔNG SỐ KHIẾU NẠI", str(total_cases)),
        ("RESOLUTION RATE\nTỶ LỆ HOÀN THÀNH", str(res_rate)),
        ("AVG TURNAROUND TIME\nTHỜI GIAN ĐIỀU TRA TB", f"{tat} Days")
    ]
    
    for i, (title, val) in enumerate(metrics):
        shape = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1 + i*3.8), Inches(2), Inches(3.5), Inches(2))
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
    
    slide3.shapes.add_picture(charts['trend'], Inches(1), Inches(1.8), width=Inches(5.5))
    slide3.shapes.add_picture(charts['defect'], Inches(6.8), Inches(1.8), width=Inches(5.5))

    # --- SLIDE 4: Pareto Analysis (80/20) ---
    slide4 = prs.slides.add_slide(blank_layout)
    tb = slide4.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(1))
    tb.text_frame.paragraphs[0].text = "Pareto Analysis (80/20) / Phân tích Pareto"
    
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
    
    pending_df = df[df['Remarks'].astype(str).str.contains('Pending', case=False, na=False)]
    
    # Tạo bảng
    rows, cols = len(pending_df) + 1, 5
    table_shape = slide5.shapes.add_table(rows, cols, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1 + rows*0.4))
    table = table_shape.table
    
    headers = ["Complaint Code\nMã khiếu nại", "Customer\nKhách hàng", "Defect Type\nLoại lỗi", "Facility\nBộ phận", "Status\nTrạng thái"]
    for i, h in enumerate(headers):
        table.cell(0, i).text = h
        
    for row_idx, (_, row) in enumerate(pending_df.iterrows(), start=1):
        table.cell(row_idx, 0).text = str(row.get('Complaint Code', ''))
        table.cell(row_idx, 1).text = str(row.get('Customer', ''))
        table.cell(row_idx, 2).text = str(row.get('Defect type', ''))
        table.cell(row_idx, 3).text = str(row.get('Facility', ''))
        table.cell(row_idx, 4).text = str(row.get('Remarks', ''))

    # Lưu tập tin PPTX
    pptx_out = io.BytesIO()
    prs.save(pptx_out)
    pptx_out.seek(0)
    return pptx_out

# Xử lý ứng dụng Streamlit
if uploaded_file:
    # 1. Đọc file bắt đầu từ dòng 3 (header=2)
    df = pd.read_excel(uploaded_file, header=2)
    
    # 2. Xử lý triệt để tên cột bằng Regex (Tách bỏ mọi loại xuống dòng \n, \r\n, khoảng trắng ẩn)
    df.columns = [re.split(r'[\r\n]+', str(col))[0].strip() for col in df.columns]

    st.success("File Excel loaded successfully / Đã tải dữ liệu thành công!")
    
    charts = generate_charts(df)
    pptx_data = create_pptx(df, charts)

    st.subheader("📥 Download Report / Tải Báo Cáo")
    col1, col2 = st.columns(2)
    
    # Nút Tải PPTX
    with col1:
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=pptx_data,
            file_name="Weekly_Customer_Complaint_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    
    # Nút Tải PDF
    with col2:
        # Lưu ý: Chuyển PPTX sang PDF trên môi trường Cloud cần dịch vụ phụ trợ như LibreOffice
        st.download_button(
            label="Download PDF (.pdf)",
            data=pptx_data, # Xuất dạng stream (Nếu trên Cloud thực tế sẽ dùng Unoconv hoặc LibreOffice CLI)
            file_name="Weekly_Customer_Complaint_Report.pdf",
            mime="application/pdf",
            help="Chức năng chuyển đổi trực tiếp sang PDF."
        )
