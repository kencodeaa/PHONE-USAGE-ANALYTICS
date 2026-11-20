"""Mini App Nhóm B: PHONE USAGE ANALYTICS – Màn hình thông minh, giấc ngủ an lành.

Cấu trúc code:
1. Import và cấu hình trang + CSS giao diện tối
2. Quản lý trạng thái dữ liệu (Session State)
3. Hàm đọc file Excel/CSV
4. Thanh tiêu đề (Header) & Nút Khôi phục gốc / Xuất Excel
5. Các Tab chức năng:
   - Tab 1: Dashboard Tổng Quan (5 thẻ KPI, 3 thẻ nhóm, 2 biểu đồ Plotly)
   - Tab 2: Tự Kiểm Tra Thói Quen (Tính giờ ngủ cá nhân + Thêm vào bảng lớp)
   - Tab 3: Bảng Dữ Liệu Khảo Sát (Upload Excel/CSV, lọc, thêm/xóa bản ghi)
   - Tab 4: Tiêu Chí & Điều Kiện
   - Tab 5: Khu Vực Công Thức Excel
   - Tab 6: Báo Cáo & Kiến Thức An Toàn
"""
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from logic import (
    REQUIRED_COLUMNS,
    calculate_sleep_hours,
    create_message,
    export_excel_workbook,
    get_default_survey_data,
    prepare_dataframe,
    summarize_dataframe,
    validate_dataframe,
)

# [1] CẤU HÌNH TRANG & CSS TÙY CHỈNH
st.set_page_config(
    page_title="PHONE USAGE ANALYTICS | Màn hình thông minh – Giấc ngủ an lành",
    page_icon="📱",
    layout="wide",
)

st.markdown(
    """
    <style>
    .stApp {
        background-color: #0B0F19;
        color: #F3F4F6;
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    /* Header Banner */
    .top-header {
        display: flex;
        align-items: center;
        gap: 14px;
        margin-bottom: 0.5rem;
    }
    .app-logo {
        width: 44px;
        height: 44px;
        background: linear-gradient(135deg, #3B82F6, #2563EB);
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 22px;
        box-shadow: 0 4px 14px rgba(59, 130, 246, 0.35);
    }
    .app-title {
        font-size: 1.35rem;
        font-weight: 800;
        letter-spacing: 0.03em;
        color: #FFFFFF;
        margin: 0;
        line-height: 1.2;
    }
    .app-subtitle {
        font-size: 0.85rem;
        color: #9CA3AF;
        margin: 2px 0 0 0;
    }
    /* 5 Thẻ KPI */
    .kpi-card {
        background-color: #131A2A;
        border: 1px solid #1F2937;
        border-radius: 12px;
        padding: 16px 18px;
        height: 100%;
    }
    .kpi-top {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
    }
    .kpi-label {
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #9CA3AF;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 1.75rem;
        font-weight: 800;
        color: #FFFFFF;
        line-height: 1.15;
    }
    .kpi-sub {
        font-size: 0.78rem;
        margin-top: 6px;
        font-weight: 500;
    }
    .kpi-icon {
        width: 34px;
        height: 34px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
    }
    /* 3 Thẻ phân nhóm */
    .group-card {
        border-radius: 12px;
        padding: 18px 20px;
        height: 100%;
    }
    .group-1 {
        background: linear-gradient(145deg, #10253E 0%, #111927 100%);
        border: 1px solid #1D4ED8;
    }
    .group-2 {
        background: linear-gradient(145deg, #1E1B3A 0%, #111927 100%);
        border: 1px solid #4F46E5;
    }
    .group-3 {
        background: linear-gradient(145deg, #2D1520 0%, #111927 100%);
        border: 1px solid #DC2626;
    }
    .group-title {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .group-val {
        font-size: 1.85rem;
        font-weight: 800;
        color: #FFFFFF;
        margin: 6px 0 2px 0;
    }
    .group-total {
        font-size: 0.85rem;
        color: #D1D5DB;
        margin-bottom: 10px;
    }
    .group-desc {
        font-size: 0.78rem;
        padding-top: 8px;
        border-top: 1px solid rgba(255,255,255,0.08);
    }
    .chart-box {
        background-color: #131A2A;
        border: 1px solid #1F2937;
        border-radius: 12px;
        padding: 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown("""
    <style>
        /* Ẩn hoàn toàn thanh header mặc định của Streamlit */
        header[data-testid="stHeader"] {
            display: none;
        }
        /* Thu gọn khoảng trống thừa ở trên cùng */
        .block-container {
            padding-top: 1rem;
        }
    </style>
""", unsafe_allow_html=True)

# [2] KHỞI TẠO DỮ LIỆU TRONG SESSION STATE
if "survey_df" not in st.session_state:
    st.session_state.survey_df = get_default_survey_data()
if "data_source_label" not in st.session_state:
    st.session_state.data_source_label = "Dữ liệu mô phỏng mặc định – Lớp 9B"


# [3] HÀM ĐỌC FILE EXCEL / CSV
def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        xl = pd.ExcelFile(uploaded_file)
        sheet = "Du_lieu_khao_sat" if "Du_lieu_khao_sat" in xl.sheet_names else xl.sheet_names[0]
        df_try = pd.read_excel(uploaded_file, sheet_name=sheet, skiprows=2)
        if not set(REQUIRED_COLUMNS).issubset(set(str(c).strip() for c in df_try.columns)):
            df_try = pd.read_excel(uploaded_file, sheet_name=sheet)
        return df_try
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    raise ValueError("Định dạng chưa hỗ trợ. Vui lòng tải file .xlsx, .xls hoặc .csv.")


df = st.session_state.survey_df
summary = summarize_dataframe(df)

# [4] THANH TIÊU ĐỀ (HEADER) & NÚT THAO TÁC NHANH
hdr_left, hdr_right1, hdr_right2 = st.columns([6.5, 1.6, 1.6])
with hdr_left:
    st.markdown(
        f"""
        <div class="top-header">
            <div class="app-logo">📱</div>
            <div>
                <div class="app-title">PHONE USAGE ANALYTICS</div>
                <div class="app-subtitle">Khảo sát Thời gian Sử dụng Điện thoại & Giấc ngủ – Lớp 9B ({summary['total']} Học sinh) • Màn hình thông minh – Giấc ngủ an lành</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with hdr_right1:
    if st.button("↺ Khôi phục gốc", use_container_width=True):
        st.session_state.survey_df = get_default_survey_data()
        st.session_state.data_source_label = "Dữ liệu mô phỏng mặc định – Lớp 9B"
        st.rerun()
with hdr_right2:
    st.download_button(
        "📥 Xuất Excel",
        data=export_excel_workbook(df),
        file_name="Phone_Usage_Analytics_Lop_9B.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

# [5] THANH ĐIỀU HƯỚNG CÁC TAB
tab_dash, tab_self, tab_data, tab_criteria, tab_excel, tab_report = st.tabs([
    "📊 Dashboard Tổng Quan",
    "🧭 Tự Kiểm Tra Thói Quen",
    "⊞ Bảng Dữ Liệu Khảo Sát",
    "⚖️ Tiêu Chí & Điều Kiện",
    "🖩 Khu Vực Công Thức Excel",
    "📘 Báo Cáo & Kiến Thức",
])

# ==============================================================================
# TAB 1: DASHBOARD TỔNG QUAN
# ==============================================================================
with tab_dash:
    # Hàng 1: 5 thẻ KPI
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-top">
                    <div>
                        <div class="kpi-label">TỔNG HỌC SINH</div>
                        <div class="kpi-value">{summary['total']}</div>
                    </div>
                    <div class="kpi-icon" style="background:rgba(148,163,184,0.15); color:#E2E8F0;">👥</div>
                </div>
                <div class="kpi-sub" style="color:#9CA3AF;">👥 Lớp 9B khảo sát</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-top">
                    <div>
                        <div class="kpi-label">THỜI GIAN DÙNG ĐT TB</div>
                        <div class="kpi-value">{summary['avg_screen']} h</div>
                    </div>
                    <div class="kpi-icon" style="background:rgba(56,189,248,0.15); color:#38BDF8;">🕒</div>
                </div>
                <div class="kpi-sub" style="color:#9CA3AF;">Giờ / ngày / học sinh</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-top">
                    <div>
                        <div class="kpi-label">THỜI LƯỢNG NGỦ TB</div>
                        <div class="kpi-value">{summary['avg_sleep']} h</div>
                    </div>
                    <div class="kpi-icon" style="background:rgba(129,140,248,0.15); color:#818CF8;">🛏️</div>
                </div>
                <div class="kpi-sub" style="color:#FBBF24;">Khuyến nghị: ≥ 8h</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-top">
                    <div>
                        <div class="kpi-label">MÀN HÌNH TRƯỚC NGỦ</div>
                        <div class="kpi-value" style="color:#FBBF24;">{summary['before_bed_yes']} HS</div>
                    </div>
                    <div class="kpi-icon" style="background:rgba(251,191,36,0.15); color:#FBBF24;">🌙</div>
                </div>
                <div class="kpi-sub" style="color:#9CA3AF;">{summary['before_pct']}% tổng số học sinh</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k5:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-top">
                    <div>
                        <div class="kpi-label">CẦN LƯU Ý CẢNH BÁO</div>
                        <div class="kpi-value" style="color:#F87171;">{summary['warning_count']} HS</div>
                    </div>
                    <div class="kpi-icon" style="background:rgba(248,113,113,0.15); color:#F87171;">⚠️</div>
                </div>
                <div class="kpi-sub" style="color:#F87171;">Cần can thiệp/nhắc nhở</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.write("")

    # Hàng 2: 3 thẻ phân nhóm sử dụng điện thoại
    g1, g2, g3 = st.columns(3)
    with g1:
        st.markdown(
            f"""
            <div class="group-card group-1">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span class="group-title" style="color:#60A5FA;">NHÓM 1 (&lt; 2H)</span>
                    <span style="font-size:20px;">🛡️</span>
                </div>
                <div class="group-val">{summary['g1_count']} HS</div>
                <div class="group-total">Tổng: <b>{summary['g1_sum']}h</b></div>
                <div class="group-desc" style="color:#60A5FA;">Mức độ an toàn - Duy trì thói quen tốt</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with g2:
        st.markdown(
            f"""
            <div class="group-card group-2">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span class="group-title" style="color:#A5B4FC;">NHÓM 2 (2H - 4H)</span>
                    <span style="font-size:20px;">⚖️</span>
                </div>
                <div class="group-val">{summary['g2_count']} HS</div>
                <div class="group-total">Tổng: <b>{summary['g2_sum']}h</b></div>
                <div class="group-desc" style="color:#A5B4FC;">Mức độ trung bình - Cần quản lý thời gian</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with g3:
        st.markdown(
            f"""
            <div class="group-card group-3">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span class="group-title" style="color:#F87171;">NHÓM 3 (&gt; 4H)</span>
                    <span style="font-size:20px;">🚨</span>
                </div>
                <div class="group-val">{summary['g3_count']} HS</div>
                <div class="group-total">Tổng: <b>{summary['g3_sum']}h</b></div>
                <div class="group-desc" style="color:#F87171;">Nguy cơ cao - Cần nhắc nhở và điều chỉnh</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.write("")

    # Hàng 3: 2 Biểu đồ Plotly (Cột & Donut)
    c_left, c_right = st.columns(2)
    with c_left:
        st.markdown("##### 📊 Phân Phối Học Sinh Theo Nhóm Sử Dụng Điện Thoại")
        fig_bar = go.Figure(
            data=[
                go.Bar(
                    x=["Nhóm 1 (< 2h)", "Nhóm 2 (2h – 4h)", "Nhóm 3 (> 4h)"],
                    y=[summary["g1_count"], summary["g2_count"], summary["g3_count"]],
                    marker_color=["#3B82F6", "#60A5FA", "#EF4444"],
                    text=[f"{summary['g1_count']} HS", f"{summary['g2_count']} HS", f"{summary['g3_count']} HS"],
                    textposition="outside",
                )
            ]
        )
        fig_bar.update_layout(
            paper_bgcolor="#131A2A",
            plot_bgcolor="#131A2A",
            font=dict(color="#E5E7EB"),
            margin=dict(l=20, r=20, t=25, b=20),
            height=330,
            yaxis=dict(gridcolor="#1F2937", title="Số học sinh", dtick=1),
            xaxis=dict(gridcolor="#1F2937"),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with c_right:
        st.markdown("##### 🌙 Thói Quen Dùng Điện Thoại 1 Giờ Trước Khi Ngủ")
        fig_donut = go.Figure(
            data=[
                go.Pie(
                    labels=["Có dùng trước ngủ", "Không dùng trước ngủ"],
                    values=[summary["before_bed_yes"], summary["before_bed_no"]],
                    hole=0.58,
                    marker=dict(colors=["#F59E0B", "#3B82F6"], line=dict(color="#131A2A", width=3)),
                    textinfo="label+percent",
                    textfont=dict(color="#F3F4F6"),
                )
            ]
        )
        fig_donut.update_layout(
            paper_bgcolor="#131A2A",
            plot_bgcolor="#131A2A",
            font=dict(color="#E5E7EB"),
            margin=dict(l=20, r=20, t=25, b=20),
            height=330,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.12, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_donut, use_container_width=True)


# ==============================================================================
# TAB 2: TỰ KIỂM TRA THÓI QUEN
# ==============================================================================
with tab_self:
    st.subheader("🧭 Tự xem lại một số thói quen cá nhân")
    st.write("Nhập thông tin để xem cách ứng dụng phân tích. Không cần nhập họ tên hay thông tin định danh.")

    c1, c2 = st.columns(2)
    with c1:
        screen_hours = st.number_input(
            "Thời gian sử dụng màn hình (giờ/ngày)",
            min_value=0.0,
            max_value=24.0,
            value=3.0,
            step=0.5,
        )
        before_sleep = st.radio(
            "Có sử dụng màn hình trong 1 giờ trước khi ngủ?",
            ["Không", "Có"],
            horizontal=True,
        )
    with c2:
        bedtime = st.time_input("Giờ đi ngủ", value=pd.Timestamp("2026-01-01 22:00").time())
        wake_time = st.time_input("Giờ thức dậy", value=pd.Timestamp("2026-01-02 06:30").time())

    sleep_hours = calculate_sleep_hours(bedtime, wake_time)
    if sleep_hours is None:
        st.warning("Không tính được thời lượng ngủ. Hãy kiểm tra lại giờ đi ngủ và giờ thức dậy.")
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("Thời lượng ngủ tính được", f"{sleep_hours:.1f} giờ")
        m2.metric(
            "Nhóm thời lượng ngủ",
            "Dưới 8 giờ" if sleep_hours < 8 else ("8–10 giờ" if sleep_hours <= 10 else "Trên 10 giờ"),
        )
        m3.metric(
            "Phân nhóm dùng điện thoại",
            "Nhóm 1 (< 2h)" if screen_hours < 2 else ("Nhóm 2 (2h – 4h)" if screen_hours <= 4 else "Nhóm 3 (> 4h)"),
        )

        msg = create_message(screen_hours, before_sleep, sleep_hours)
        if msg.startswith("Cần"):
            st.warning(f"⚠️ {msg}")
        else:
            st.success(f"✅ {msg}")

        if st.button("➕ Thêm kết quả ẩn danh này vào dữ liệu khảo sát lớp 9B"):
            next_num = len(st.session_state.survey_df) + 1
            new_row = pd.DataFrame([{
                "Mã HS": f"HS{next_num:02d}",
                "Thời gian sử dụng màn hình (giờ/ngày)": float(screen_hours),
                "Có sử dụng màn hình trong 1 giờ trước khi ngủ? (Có/Không)": before_sleep,
                "Giờ đi ngủ": bedtime.strftime("%H:%M"),
                "Giờ thức dậy": wake_time.strftime("%H:%M"),
                "Thời lượng ngủ (giờ)": sleep_hours,
                "Tự đánh giá ngủ đủ? (Có/Không)": "Có" if sleep_hours >= 8 else "Không",
            }])
            st.session_state.survey_df = prepare_dataframe(
                pd.concat([st.session_state.survey_df, new_row], ignore_index=True)
            )
            st.success(f"Đã thêm mã HS{next_num:02d} vào bảng khảo sát! Chuyển sang Dashboard để xem cập nhật.")
            st.rerun()

    st.caption("Ngưỡng >4 giờ/ngày là quy ước học tập của dự án; mốc 8–10 giờ là khuyến nghị tham khảo phổ biến cho thanh thiếu niên 13–18 tuổi.")


# ==============================================================================
# TAB 3: BẢNG DỮ LIỆU KHẢO SÁT
# ==============================================================================
with tab_data:
    st.subheader("⊞ Bảng dữ liệu khảo sát lớp 9B")
    st.caption(f"Nguồn dữ liệu hiện tại: **{st.session_state.data_source_label}**")

    up_col1, up_col2 = st.columns([2, 1])
    with up_col1:
        uploaded = st.file_uploader(
            "Tải file Excel Nhóm B (.xlsx, .xls) hoặc CSV ẩn danh để thay thế dữ liệu phân tích",
            type=["xlsx", "xls", "csv"],
            key="survey_upload",
        )
        if uploaded is not None:
            try:
                raw = read_uploaded_file(uploaded)
                prepared = prepare_dataframe(raw)
                st.session_state.survey_df = prepared
                st.session_state.data_source_label = f"File tải lên: {uploaded.name}"
                df = prepared
                issues = validate_dataframe(df)
                if issues:
                    st.warning("Cần rà soát dữ liệu trước khi sử dụng kết quả:")
                    for issue in issues:
                        st.write(f"- {issue}")
                else:
                    st.success(f"Đã nạp thành công {len(df)} bản ghi hợp lệ từ {uploaded.name}.")
            except Exception as exc:
                st.error(f"Không thể xử lý file: {exc}")

    with up_col2:
        st.markdown("**Tải biểu mẫu CSV chuẩn**")
        sample_csv = df[REQUIRED_COLUMNS].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Tải file CSV mẫu (30 HS)",
            data=sample_csv,
            file_name="du_lieu_mo_phong_30_HS.csv",
            mime="text/csv",
            use_container_width=True,
        )

    f1, f2 = st.columns([2, 1])
    with f1:
        search_id = st.text_input("🔍 Tìm kiếm theo Mã HS (ví dụ: HS05)", "")
    with f2:
        group_filter = st.selectbox(
            "Lọc theo nhóm",
            ["Tất cả", "Nhóm 1 (< 2h)", "Nhóm 2 (2h – 4h)", "Nhóm 3 (> 4h)", "Cần lưu ý cảnh báo"],
        )

    view_df = df.copy()
    if search_id.strip():
        view_df = view_df[view_df["Mã HS"].str.contains(search_id.strip(), case=False, na=False)]
    if group_filter == "Cần lưu ý cảnh báo":
        view_df = view_df[view_df["Thông điệp gợi ý"] == "Cần xem lại thói quen màn hình và giấc ngủ"]
    elif group_filter != "Tất cả":
        view_df = view_df[view_df["Phân nhóm ĐT"] == group_filter]

    st.dataframe(view_df, use_container_width=True, hide_index=True)


# ==============================================================================
# TAB 4: TIÊU CHÍ & ĐIỀU KIỆN
# ==============================================================================
with tab_criteria:
    st.subheader("⚖️ Quy ước phân nhóm & tiêu chí đánh giá STEM Tin học 9")
    c_a, c_b = st.columns(2)
    with c_a:
        st.info(
            "**1. Phân loại thời gian sử dụng điện thoại/màn hình**\n"
            "- **Nhóm 1 (< 2 giờ/ngày):** Mức độ an toàn – Duy trì thói quen tốt.\n"
            "- **Nhóm 2 (2 – 4 giờ/ngày):** Mức độ trung bình – Cần quản lý thời gian hợp lý.\n"
            "- **Nhóm 3 (> 4 giờ/ngày):** Ngưỡng cảnh báo theo quy ước học tập của dự án."
        )
    with c_b:
        st.info(
            "**2. Phân loại thời lượng ngủ & thói quen trước ngủ**\n"
            "- **Dưới 8 giờ:** Nhóm dưới mốc khuyến nghị tham khảo cho tuổi 13–18.\n"
            "- **8 – 10 giờ:** Mốc khuyến nghị tham khảo phổ biến cho thanh thiếu niên.\n"
            "- **Dùng màn hình 1 giờ trước ngủ = Có:** Gắn nhãn cần nhắc nhở điều chỉnh thói quen."
        )

    st.markdown("#### Bảng tổng hợp các chỉ tiêu theo quy ước")
    stat_table = pd.DataFrame([
        ["Tổng số học sinh khảo sát", summary["total"], "Đếm số mã HS hợp lệ"],
        ["Nhóm 1 (< 2 giờ/ngày)", f"{summary['g1_count']} HS ({summary['g1_sum']}h)", "Mức độ an toàn"],
        ["Nhóm 2 (2 – 4 giờ/ngày)", f"{summary['g2_count']} HS ({summary['g2_sum']}h)", "Mức độ trung bình"],
        ["Nhóm 3 (> 4 giờ/ngày)", f"{summary['g3_count']} HS ({summary['g3_sum']}h)", "Tương đương COUNTIF / SUMIF > 4"],
        ["Số HS ngủ dưới 8 giờ", f"{summary['sleep_under']} HS ({summary['sum_sleep_under']}h)", "Tương đương COUNTIF / SUMIF < 8"],
        ["Số HS dùng màn hình trong 1h trước ngủ", f"{summary['before_bed_yes']} HS ({summary['before_pct']}%)", "Tương đương COUNTIF = 'Có'"],
        ["Số HS cần lưu ý cảnh báo chung", f"{summary['warning_count']} HS", "Thỏa mãn ít nhất 1 trong 3 điều kiện rủi ro"],
    ], columns=["Chỉ tiêu", "Kết quả hiện tại", "Ghi chú quy ước"])
    st.dataframe(stat_table, use_container_width=True, hide_index=True)


# ==============================================================================
# TAB 5: KHU VỰC CÔNG THỨC EXCEL
# ==============================================================================
with tab_excel:
    st.subheader("🖩 Khu vực công thức Excel (Đối chiếu chương trình Tin học 9)")
    st.write("Các phép tính trong Python của ứng dụng này mô phỏng trực tiếp các hàm bảng tính dưới đây (giả sử dữ liệu từ dòng 4 đến dòng 33):")

    formulas = [
        ("1. Tính thời lượng ngủ qua nửa đêm (cột F)", '=ROUND(MOD(E4 - D4, 1) * 24, 1)', "Hàm MOD xử lý trường hợp đi ngủ trước 24:00 (VD 22:30) và thức dậy sau 00:00 (VD 06:30)."),
        ("2. Phân loại 3 nhóm sử dụng điện thoại", '=IF(B4<2, "Nhóm 1 (< 2h)", IF(B4<=4, "Nhóm 2 (2h - 4h)", "Nhóm 3 (> 4h)"))', "Hàm IF lồng nhau phân loại mức độ sử dụng màn hình theo ngày."),
        ("3. Thông điệp cảnh báo tổng hợp (IF + OR)", '=IF(OR(B4>4, C4="Có", F4<8), "Cần xem lại thói quen màn hình và giấc ngủ", "Thói quen tương đối phù hợp")', "Hàm OR kết hợp 3 điều kiện: màn hình > 4h HOẶC dùng trước ngủ HOẶC ngủ < 8h."),
        ("4. Đếm số học sinh Nhóm 3 (> 4h) – COUNTIF", '=COUNTIF(B4:B33, ">4")', "Đếm có điều kiện các học sinh vượt ngưỡng 4 giờ/ngày."),
        ("5. Tính tổng giờ màn hình Nhóm 3 – SUMIF", '=SUMIF(B4:B33, ">4", B4:B33)', "Tính tổng có điều kiện số giờ sử dụng của nhóm > 4 giờ/ngày."),
        ("6. Đếm số học sinh ngủ dưới 8 giờ – COUNTIF", '=COUNTIF(F4:F33, "<8")', "Đếm số học sinh có thời lượng ngủ chưa đạt khuyến nghị 8 giờ."),
    ]
    for title, formula, desc in formulas:
        with st.container(border=True):
            st.markdown(f"**{title}**")
            st.code(formula, language="excel")
            st.caption(desc)


# ==============================================================================
# TAB 6: BÁO CÁO & KIẾN THỨC AN TOÀN
# ==============================================================================
with tab_report:
    st.subheader("📘 Kiến thức nền, Đạo đức dữ liệu & Báo cáo lớp 9B")
    st.markdown(
        "**Một số điểm cần nhớ**\n"
        "- Thanh thiếu niên 13–18 tuổi thường được khuyến nghị ngủ khoảng 8–10 giờ mỗi ngày.\n"
        "- Sử dụng màn hình sát giờ ngủ có thể làm giảm thời gian dành cho nghỉ ngơi hoặc gây khó thư giãn.\n"
        "- Khảo sát tự báo cáo có thể có sai số; dữ liệu lớp học nhỏ không đại diện cho tất cả học sinh."
    )
    st.warning(
        "Không dùng ứng dụng để chẩn đoán, xếp loại sức khỏe hoặc so sánh công khai từng học sinh. "
        "Nếu khó ngủ/mệt mỏi kéo dài, hãy trao đổi với cha mẹ/người chăm sóc hoặc nhân viên y tế."
    )
    st.markdown("**Vai trò của AI trong dự án**")
    st.write(
        "AI chỉ hỗ trợ học sinh và giáo viên trong giai đoạn thiết kế cấu trúc, tạo bản nháp code và giải thích lỗi. "
        "Khi ứng dụng đang chạy, mọi phép tính do Python thực hiện theo quy tắc lập trình sẵn; không có kết nối API AI trả phí."
    )
    st.markdown("#### Tóm tắt báo cáo nhanh (Sẵn sàng sao chép / in ấn)")
    report_text = (
        f"BÁO CÁO KHẢO SÁT THỜI GIAN SỬ DỤNG ĐIỆN THOẠI & GIẤC NGỦ – LỚP 9B\n"
        f"- Tổng số học sinh tham gia (ẩn danh): {summary['total']} học sinh\n"
        f"- Thời gian dùng điện thoại trung bình: {summary['avg_screen']} giờ/ngày\n"
        f"- Thời lượng ngủ trung bình: {summary['avg_sleep']} giờ/đêm (Khuyến nghị: 8–10 giờ)\n"
        f"- Phân bố nhóm sử dụng ĐT: Nhóm 1 (<2h): {summary['g1_count']} HS | Nhóm 2 (2-4h): {summary['g2_count']} HS | Nhóm 3 (>4h): {summary['g3_count']} HS\n"
        f"- Tỷ lệ dùng màn hình trong 1 giờ trước khi ngủ: {summary['before_bed_yes']}/{summary['total']} HS ({summary['before_pct']}%)\n"
        f"- Số trường hợp cần nhắc nhở điều chỉnh thói quen: {summary['warning_count']} HS\n"
        f"Ghi chú: Số liệu phục vụ học tập STEM Tin học 9, không mang tính chẩn đoán y tế."
    )
    st.text_area("Nội dung báo cáo tổng hợp", value=report_text, height=200)

st.divider()
st.caption("Sản phẩm học tập STEM – Tin học 9 | Nhóm B | Dữ liệu vì sức khỏe học đường")
