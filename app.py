from datetime import datetime
import io
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Return Dashboard", layout="wide")

# ======================
# HÀM XUẤT EXCEL TỰ TÍCH HỢP (KHÔNG CẦN FILE UTILS)
# ======================
def convert_df_to_excel(df_to_export):
    output = io.BytesIO()
    # Sử dụng xlsxwriter để tạo file Excel chuẩn hóa
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_to_export.to_excel(writer, index=False, sheet_name="Return_Report")
    return output.getvalue()

# ======================
# LOAD DATA (TẢI FILE TRỰC TIẾP TỪ TRÌNH DUYỆT)
# ======================
st.sidebar.header("DỮ LIỆU ĐẦU VÀO")

uploaded_file = st.sidebar.file_uploader(
    "Tải lên file Excel dữ liệu (Return List)",
    type=["xlsx"]
)

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, header=1)
    df.columns = [c.strip() for c in df.columns]
    df["Date"] = pd.to_datetime(df["Date"])
    return df

if uploaded_file is not None:
    try:
        df = load_data(uploaded_file)
    except Exception as e:
        st.error(f"❌ Cấu trúc file không đúng hoặc bị lỗi định dạng: {e}")
        st.stop()
else:
    st.info("👋 Vui lòng tải file Excel dữ liệu lên ở Sidebar!")
    st.stop()

# ======================
# SIDEBAR FILTER
# ======================
st.sidebar.header("FILTER")

min_date = df["Date"].min() if not df.empty else datetime.today()
max_date = df["Date"].max() if not df.empty else datetime.today()

date_range = st.sidebar.date_input("Date Range", [min_date, max_date])
customer = st.sidebar.multiselect("Customer", df["Customer"].unique())
facility = st.sidebar.multiselect("Facility", df["Facility"].unique())
defect_type = st.sidebar.multiselect("Defect Type", df["Defect type"].unique())

# ======================
# FILTER DATA
# ======================
filtered = df.copy()

if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    start_date, end_date = date_range
    filtered = filtered[
        (filtered["Date"] >= pd.to_datetime(start_date)) &
        (filtered["Date"] <= pd.to_datetime(end_date))
    ]

if customer:
    filtered = filtered[filtered["Customer"].isin(customer)]

if facility:
    filtered = filtered[filtered["Facility"].isin(facility)]

if defect_type:
    filtered = filtered[filtered["Defect type"].isin(defect_type)]

if filtered.empty:
    st.warning("⚠️ Không tìm thấy dữ liệu phù hợp!")
    st.stop()

# ======================
# KPI
# ======================
col1, col2, col3, col4 = st.columns(4)

col1.metric("Total Return", int(filtered["Return Quantity"].sum()))
col2.metric("Total Defect", int(filtered["Defect"].sum()))

avg_defect = filtered["Percentage"].mean()
col3.metric("Avg % Defect", f"{round(avg_defect, 2)}%" if not pd.isna(avg_defect) else "0%")

col4.metric("Customers", filtered["Customer"].nunique())

st.divider()

# ======================
# CHARTS
# ======================
c1, c2 = st.columns(2)

with c1:
    fig = px.bar(
        filtered,
        x="Customer",
        y="Defect",
        color="Facility",
        title="Defect by Customer",
    )

    fig.update_layout(
        font=dict(color="hotpink"),
        title_font_color="hotpink",
        legend_title_font_color="hotpink"
    )

    fig.update_xaxes(title_font=dict(color="hotpink"), tickfont=dict(color="hotpink"))
    fig.update_yaxes(title_font=dict(color="hotpink"), tickfont=dict(color="hotpink"))

    st.plotly_chart(fig, use_container_width=True)

with c2:
    fig2 = px.pie(
        filtered,
        names="Defect type",
        values="Defect",
        title="Defect Type Distribution",
    )

    fig2.update_layout(
        font=dict(color="hotpink"),
        title_font_color="hotpink",
        legend_title_font_color="hotpink"
    )

    # pie chart không có x/y axis nhưng vẫn giữ consistency
    st.plotly_chart(fig2, use_container_width=True)

trend_df = filtered.groupby("Date")["Defect"].sum().reset_index()

fig3 = px.line(
    trend_df,
    x="Date",
    y="Defect",
    title="Defect Trend"
)

fig3.update_layout(
    font=dict(color="hotpink"),
    title_font_color="hotpink",
    legend_title_font_color="hotpink"
)

fig3.update_xaxes(title_font=dict(color="hotpink"), tickfont=dict(color="hotpink"))
fig3.update_yaxes(title_font=dict(color="hotpink"), tickfont=dict(color="hotpink"))

st.plotly_chart(fig3, use_container_width=True)

# ======================
# TABLE
# ======================
st.subheader("DATA DETAIL")
st.dataframe(filtered.sort_values(by="Defect", ascending=False), use_container_width=True)

# ======================
# EXPORT
# ======================
st.subheader("EXPORT REPORT")

excel_data = convert_df_to_excel(filtered)

st.download_button(
    label="⬇️ Download Excel Report",
    data=excel_data,
    file_name="return_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
