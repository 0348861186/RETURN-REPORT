import io
import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# Cấu hình giao diện trang web
st.set_page_config(
    page_title="Weekly CCR Dashboard", page_icon="📊", layout="wide"
)

st.title("📊 Weekly Customer Complaint Dashboard & PPT Generator")
st.markdown(
    "Ứng dụng tổng hợp dữ liệu khiếu nại, hiển thị Dashboard trực quan và xuất"
    " báo cáo PowerPoint tự động."
)

# ============================================================
# XỬ LÝ DỮ LIỆU & PHÒNG NGỪA LỖI KEYERROR: 'Date'
# ============================================================
uploaded_file = st.file_uploader(
    "Tải lên file dữ liệu khiếu nại (Excel/CSV có chứa cột Date hoặc Ngày)",
    type=["csv", "xlsx"],
)

if uploaded_file is not None:
  try:
    if uploaded_file.name.endswith(".csv"):
      df = pd.read_csv(uploaded_file)
    else:
      df = pd.read_excel(uploaded_file)

    # Khắc phục lỗi KeyError: 'Date' bằng cách kiểm tra linh hoạt tên cột
    if "Date" not in df.columns:
      # Tìm các cột có tên gần giống như 'date', 'ngay', 'Date', 'Ngày'
      possible_cols = [
          c
          for c in df.columns
          if c.lower() in ["date", "ngày", "ngay_thang", "time"]
      ]
      if possible_cols:
        df = df.rename(columns={possible_cols[0]: "Date"})
        st.warning(
            f"⚠️ Đã tự động nhận diện cột '{possible_cols[0]}' làm cột 'Date'."
        )
      else:
        st.error(
            "❌ Lỗi: File dữ liệu thiếu cột **'Date'** (hoặc 'Ngày'). Vui lòng"
            " kiểm tra lại cấu trúc file!"
        )
        st.stop()
  except Exception as e:
    st.error(f"❌ Lỗi khi đọc file: {e}")
    st.stop()
else:
  # Dữ liệu mẫu mặc định nếu chưa upload file (giúp app không bị crash)
  st.info(
      "💡 Đang hiển thị dữ liệu mẫu. Hãy tải lên file của bạn để xem dữ liệu"
      " thực tế."
  )
  df = pd.DataFrame({
      "Date": ["2026-08-01", "2026-08-02", "2026-08-03"],
      "Complaint Code": ["CCR-2026-001", "CCR-2026-002", "CCR-2026-003"],
      "Facility": ["Plant A", "Plant B", "Plant A"],
      "Customer": ["Customer A", "Customer B", "Customer C"],
      "Status": ["OPEN", "CLOSED", "OPEN"],
  })

# Hiển thị bảng dữ liệu lên Dashboard
st.subheader("📋 Bảng dữ liệu chi tiết")
st.dataframe(df, use_container_width=True)


# ============================================================
# HÀM TẠO POWERPOINT TRONG BỘ NHỚ (BYTES IO)
# ============================================================
def generate_pptx_stream(data_df):
  prs = Presentation()
  prs.slide_width = Inches(13.333)
  prs.slide_height = Inches(7.5)
  blank_layout = prs.slide_layouts[6]

  COLOR_BG = RGBColor(245, 247, 250)
  COLOR_BLUE = RGBColor(37, 99, 235)
  COLOR_GREEN = RGBColor(16, 185, 129)
  COLOR_ORANGE = RGBColor(245, 158, 11)
  COLOR_RED = RGBColor(220, 38, 38)
  COLOR_DARK = RGBColor(30, 41, 59)
  COLOR_WHITE = RGBColor(255, 255, 255)
  COLOR_GRAY = RGBColor(100, 116, 139)

  def set_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_BG

  def add_title_banner(slide, title_en, title_vi, color):
    banner = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5),
        Inches(0.4),
        Inches(12.333),
        Inches(0.8),
    )
    banner.fill.solid()
    banner.fill.fore_color.rgb = color
    banner.line.color.rgb = color
    tf = banner.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_en
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = title_vi
    p2.font.size = Pt(11)
    p2.font.color.rgb = COLOR_WHITE
    p2.alignment = PP_ALIGN.CENTER

  # SLIDE 1 - COVER
  slide = prs.slides.add_slide(blank_layout)
  set_background(slide)
  title = slide.shapes.add_shape(
      MSO_SHAPE.ROUNDED_RECTANGLE,
      Inches(1.0),
      Inches(1.7),
      Inches(11.3),
      Inches(2.2),
  )
  title.fill.solid()
  title.fill.fore_color.rgb = COLOR_BLUE
  title.line.color.rgb = COLOR_BLUE
  tf = title.text_frame
  tf.clear()

  p = tf.paragraphs[0]
  p.text = "WEEKLY CUSTOMER COMPLAINT REPORT"
  p.font.size = Pt(30)
  p.font.bold = True
  p.font.color.rgb = COLOR_WHITE
  p.alignment = PP_ALIGN.CENTER

  p2 = tf.add_paragraph()
  p2.text = "BÁO CÁO KHIẾU NẠI KHÁCH HÀNG HÀNG TUẦN"
  p2.font.size = Pt(18)
  p2.font.color.rgb = COLOR_WHITE
  p2.alignment = PP_ALIGN.CENTER

  # Xuất file ra stream trong bộ nhớ thay vì lưu ổ cứng
  file_stream = io.BytesIO()
  prs.save(file_stream)
  file_stream.seek(0)
  return file_stream


# ============================================================
# GIAO DIỆN TẢI XUỐNG POWERPOINT
# ============================================================
st.markdown("---")
st.subheader("📥 Xuất Báo Cáo Tự Động")

if st.button("🚀 Tạo và Tải File PowerPoint Báo Cáo"):
  try:
    ppt_data = generate_pptx_stream(df)
    st.success("✅ Đã tạo file PowerPoint thành công!")
    st.download_button(
        label="⬇️ Click để tải xuống file .pptx",
        data=ppt_data,
        file_name="Weekly_CCR_Report.pptx",
        mime=(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        ),
    )
  except Exception as e:
    st.error(f"❌ Lỗi khi tạo PowerPoint: {e}")
