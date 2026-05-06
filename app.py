import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import traceback

# ─── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Tạo SBD Tự Động",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Root tokens ── */
:root {
    --primary: #2563EB;
    --primary-light: #3B82F6;
    --accent: #10B981;
    --warn: #F59E0B;
    --danger: #EF4444;
    --radius: 12px;
}

/* ── Global ── */
.stApp { background: var(--bg, #F0F4FF); }

/* ── Title ── */
.app-title {
    text-align: center;
    font-size: 2rem;
    font-weight: 800;
    letter-spacing: .5px;
    padding: 1.4rem 0 .6rem;
    background: linear-gradient(135deg, #2563EB 0%, #7C3AED 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.app-subtitle {
    text-align: center;
    color: #6B7280;
    font-size: .95rem;
    margin-bottom: 1.8rem;
}

/* ── Section cards ── */
.section-card {
    background: #FFFFFF;
    border-radius: var(--radius);
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 2px 12px rgba(37,99,235,.08);
    border: 1px solid #E5E7EB;
}
.section-label {
    font-size: .78rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #2563EB;
    margin-bottom: .7rem;
}
.section-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #1E293B;
    margin-bottom: .9rem;
}

/* ── Exam radio pills ── */
div[data-testid="stRadio"] > div { gap: 10px !important; flex-wrap: wrap; }
div[data-testid="stRadio"] label {
    background: #F1F5FF !important;
    border: 2px solid #BFDBFE !important;
    border-radius: 50px !important;
    padding: .45rem 1.2rem !important;
    font-weight: 600 !important;
    color: #2563EB !important;
    cursor: pointer;
    transition: all .2s;
}
div[data-testid="stRadio"] label:has(input:checked) {
    background: #2563EB !important;
    border-color: #2563EB !important;
    color: #FFF !important;
}

/* ── Buttons ── */
.stDownloadButton > button, .stButton > button {
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: .55rem 1.4rem !important;
}
.stDownloadButton > button {
    background: #F1F5FF !important;
    border: 2px solid #BFDBFE !important;
    color: #2563EB !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2563EB, #7C3AED) !important;
    border: none !important;
    color: white !important;
}

/* ── Success / Error banner ── */
.banner-ok {
    background: #ECFDF5; border-left: 4px solid #10B981;
    padding: .9rem 1.2rem; border-radius: 8px; margin-bottom: 1rem;
    color: #065F46; font-weight: 600;
}
.banner-err {
    background: #FFF5F5; border-left: 4px solid #EF4444;
    padding: .9rem 1.2rem; border-radius: 8px; margin-bottom: 1rem;
    color: #991B1B; font-weight: 600;
}

/* ── Dark mode overrides ── */
@media (prefers-color-scheme: dark) {
    .stApp { background: #0F172A !important; }
    .section-card { background: #1E293B !important; border-color: #334155 !important; }
    .section-title { color: #F1F5F9 !important; }
    .app-subtitle { color: #94A3B8 !important; }
    div[data-testid="stRadio"] label { background: #1E3A5F !important; border-color: #3B82F6 !important; color: #93C5FD !important; }
    div[data-testid="stRadio"] label:has(input:checked) { background: #2563EB !important; color: #FFF !important; }
}

div[data-testid="stFileUploader"] {
    border: 2px dashed #BFDBFE;
    border-radius: var(--radius);
    padding: .5rem;
}
</style>
""", unsafe_allow_html=True)


# ─── CONSTANTS ─────────────────────────────────────────────────────────────────
EXAM_RULES = {
    "BEBRAS": {
        "label": "🦫 BEBRAS",
        "desc": "Lớp 1–12 (6 cấp độ)",
        "valid_grades": list(range(1, 13)),
        "use_level": True,
    },
    "AMC8": {
        "label": "🔢 AMC8",
        "desc": "Lớp 4–8",
        "valid_grades": [4, 5, 6, 7, 8],
        "use_level": False,
    },
    "AMC1012": {
        "label": "📐 AMC10/12",
        "desc": "AMC10: Lớp 6–10 | AMC12: Lớp 11–12",
        "valid_grades": list(range(6, 13)),
        "use_level": False,
        "sub_levels": {
            "AMC10": [6, 7, 8, 9, 10],
            "AMC12": [11, 12],
        },
    },
    "VEO": {
        "label": "🌿 VEO",
        "desc": "JUNIOR: Lớp 6–9 | VEO: Lớp 10–12",
        "valid_grades": list(range(6, 13)),
        "use_level": False,
        "sub_levels": {
            "VEO JUNIOR": [6, 7, 8, 9],
            "VEO": [10, 11, 12],
        },
    },
}

SAMPLE_PATH = "SBD_mẫu.xlsx"


# ─── HELPERS ───────────────────────────────────────────────────────────────────
def load_sample_bytes():
    with open(SAMPLE_PATH, "rb") as f:
        return f.read()


def validate_input(df_hs: pd.DataFrame, df_diem: pd.DataFrame, df_xep: pd.DataFrame, exam: str):
    errors = []
    required_hs = ["STT", "ID gốc", "Họ và tên đệm", "Tên", "Ngày sinh", "Tháng sinh",
                   "Năm sinh", "Giới tính", "Khối lớp", "Lớp", "Xã/Phường", "Tỉnh thành",
                   "Cấp độ", "HS nước ngoài", "Cụm"]
    for c in required_hs:
        if c not in df_hs.columns:
            errors.append(f"Sheet 'Học-sinh' thiếu cột: **{c}**")

    required_diem = ["Điểm thi", "Phòng", "Số lượng hs/phòng"]
    for c in required_diem:
        if c not in df_diem.columns:
            errors.append(f"Sheet 'Điểm-thi' thiếu cột: **{c}**")

    required_xep = ["Cụm", "Điểm thi"]
    for c in required_xep:
        if c not in df_xep.columns:
            errors.append(f"Sheet 'Xếp-điểm-thi' thiếu cột: **{c}**")

    if errors:
        return errors

    # Check valid grades
    rules = EXAM_RULES[exam]
    invalid_grades = df_hs[~df_hs["Khối lớp"].isin(rules["valid_grades"])]["Khối lớp"].unique()
    if len(invalid_grades) > 0:
        errors.append(f"Có học sinh không thuộc khối hợp lệ cho {exam}: {list(invalid_grades)}")

    # Check all Cụm have a Điểm thi
    cum_list = df_hs["Cụm"].dropna().astype(str).str.strip().unique()
    xep_clean_v = df_xep.dropna(subset=["Cụm", "Điểm thi"]).copy()
    xep_clean_v["Cụm"] = xep_clean_v["Cụm"].astype(str).str.strip()
    xep_map = xep_clean_v.set_index("Cụm")["Điểm thi"].to_dict()
    missing_cum = [c for c in cum_list if c not in xep_map]
    if missing_cum:
        errors.append(f"Các Cụm chưa được gán Điểm thi: {missing_cum}")

    return errors


def assign_level_bebras(grade):
    if grade in [1, 2]: return 1
    if grade in [3, 4]: return 2
    if grade in [5, 6]: return 3
    if grade in [7, 8]: return 4
    if grade in [9, 10]: return 5
    if grade in [11, 12]: return 6
    return None


def assign_sub_level(exam, grade):
    rules = EXAM_RULES.get(exam, {})
    if "sub_levels" not in rules:
        return None
    for name, grades in rules["sub_levels"].items():
        if grade in grades:
            return name
    return None


def process_sbd(df_hs_raw: pd.DataFrame, df_diem: pd.DataFrame, df_xep: pd.DataFrame, exam: str):
    # ── Filter valid data ──
    df = df_hs_raw.copy()
    # Drop helper rows/cols
    df = df[df["Khối lớp"].notna() & df["Tên"].notna()].copy()
    df = df[df["Khối lớp"].isin(EXAM_RULES[exam]["valid_grades"])].copy()

    if len(df) == 0:
        raise ValueError("Không có học sinh hợp lệ sau khi lọc theo kỳ thi.")

    # Ensure types
    df["HS nước ngoài"] = pd.to_numeric(df["HS nước ngoài"], errors="coerce").fillna(0).astype(int)
    df["Khối lớp"] = pd.to_numeric(df["Khối lớp"], errors="coerce").astype(int)

    # BEBRAS: ensure Cấp độ is correct
    if exam == "BEBRAS":
        df["Cấp độ"] = df["Khối lớp"].apply(assign_level_bebras)
    elif exam in ("AMC1012", "VEO"):
        df["Cấp độ"] = df["Khối lớp"].apply(lambda g: assign_sub_level(exam, g))
    else:
        df["Cấp độ"] = df["Cấp độ"].fillna("")

    # ── Sort Phase 1 ──
    sort_cols1 = ["HS nước ngoài", "Cấp độ", "Khối lớp", "Tên", "Họ và tên đệm"]
    # Cấp độ may be string for AMC/VEO, handle numeric sort for BEBRAS
    df = df.sort_values(sort_cols1, ascending=True, na_position="last").reset_index(drop=True)
    df["STT"] = range(1, len(df) + 1)

    # ── Build Điểm thi map from Xếp-điểm-thi ──
    xep_clean = df_xep.dropna(subset=["Cụm", "Điểm thi"]).copy()
    xep_clean["Cụm"] = xep_clean["Cụm"].astype(str).str.strip()
    xep_map = xep_clean.set_index("Cụm")["Điểm thi"].to_dict()
    df["Cụm"] = df["Cụm"].astype(str).str.strip()
    df["Điểm thi"] = df["Cụm"].map(xep_map)

    # ── Build room capacity map ──
    diem_clean = df_diem.dropna(subset=["Điểm thi", "Phòng", "Số lượng hs/phòng"])
    # Convert to numeric
    diem_clean = diem_clean.copy()
    diem_clean["Điểm thi"] = pd.to_numeric(diem_clean["Điểm thi"], errors="coerce")
    diem_clean["Phòng"] = pd.to_numeric(diem_clean["Phòng"], errors="coerce")
    diem_clean["Số lượng hs/phòng"] = pd.to_numeric(diem_clean["Số lượng hs/phòng"], errors="coerce")
    diem_clean = diem_clean.dropna()

    # Group by Điểm thi → list of (Phòng, capacity)
    room_map = {}
    for dt, grp in diem_clean.groupby("Điểm thi"):
        rooms = sorted(grp[["Phòng", "Số lượng hs/phòng"]].values.tolist(), key=lambda x: x[0])
        room_map[dt] = rooms

    # ── Assign rooms ──
    phong_col = []
    stt_phong_col = []

    # Track current room index & count per điểm thi
    room_state = {}  # dt -> (room_idx, count_in_room)
    for _, row in df.iterrows():
        dt = row["Điểm thi"]
        if pd.isna(dt):
            phong_col.append(None)
            stt_phong_col.append(None)
            continue

        dt = int(dt)
        if dt not in room_map or len(room_map[dt]) == 0:
            phong_col.append(None)
            stt_phong_col.append(None)
            continue

        if dt not in room_state:
            room_state[dt] = [0, 0]  # [room_idx, count]

        state = room_state[dt]
        rooms = room_map[dt]
        room_idx, count = state

        if room_idx >= len(rooms):
            # No more rooms – overflow
            phong_col.append(rooms[-1][0])
            stt_phong_col.append(count + 1)
            room_state[dt][1] += 1
            continue

        cap = int(rooms[room_idx][1])
        if count >= cap:
            room_idx += 1
            count = 0
            room_state[dt] = [room_idx, count]
            if room_idx >= len(rooms):
                phong_col.append(rooms[-1][0])
                stt_phong_col.append(count + 1)
                room_state[dt][1] += 1
                continue

        phong_col.append(int(rooms[room_idx][0]))
        stt_phong_col.append(count + 1)
        room_state[dt][1] = count + 1

    df["Phòng thi"] = phong_col
    df["STT trong phòng"] = stt_phong_col

    # ── SBD formula (calculate in Python, store as zero-padded string) ──
    def make_sbd(row):
        try:
            return f"{int(row['Điểm thi']):02d}{int(row['Phòng thi']):02d}{int(row['STT trong phòng']):02d}"
        except:
            return None

    df["SBD"] = df.apply(make_sbd, axis=1)

    # ── Sort Phase 2 ──
    df = df.sort_values(["Điểm thi", "Phòng thi", "STT trong phòng"],
                        ascending=True, na_position="last").reset_index(drop=True)
    df["STT"] = range(1, len(df) + 1)

    # ── Column order for SBD sheet ──
    final_cols = ["STT", "Điểm thi", "Phòng thi", "STT trong phòng", "SBD",
                  "ID gốc", "Họ và tên đệm", "Tên", "Ngày sinh", "Tháng sinh",
                  "Năm sinh", "Giới tính", "Khối lớp", "Lớp", "Xã/Phường",
                  "Tỉnh thành", "Cấp độ", "HS nước ngoài"]
    df = df[[c for c in final_cols if c in df.columns]]
    return df


def export_excel(df_sbd: pd.DataFrame, df_diem_raw, df_xep_raw):
    wb = openpyxl.Workbook()

    # ── Helper styles ──
    header_fill = PatternFill("solid", fgColor="2563EB")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )
    sbd_fill = PatternFill("solid", fgColor="EFF6FF")
    sbd_font = Font(bold=True, color="1D4ED8", name="Arial Narrow", size=11)

    def write_df_sheet(ws, df, freeze_row=2):
        ws.freeze_panes = f"A{freeze_row}"
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = thin

        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            for c_idx, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin
                cell.alignment = Alignment(vertical="center", horizontal="center" if c_idx <= 5 else "left")
                if c_idx == 5:  # SBD col
                    cell.fill = sbd_fill
                    cell.font = sbd_font

        # Auto width
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 30)
        ws.row_dimensions[1].height = 30

    # ── Sheet SBD ──
    ws_sbd = wb.active
    ws_sbd.title = "SBD"
    write_df_sheet(ws_sbd, df_sbd)

    # ── Sheet Điểm-thi ──
    ws_dt = wb.create_sheet("Điểm-thi")
    diem_cols = ["Điểm thi", "Phòng", "Số lượng hs/phòng"]
    df_dt = df_diem_raw[[c for c in diem_cols if c in df_diem_raw.columns]].dropna(subset=["Điểm thi"])
    df_dt = df_dt[pd.to_numeric(df_dt["Điểm thi"], errors="coerce").notna()]
    write_df_sheet(ws_dt, df_dt)

    # ── Sheet Xếp-điểm-thi ──
    ws_xep = wb.create_sheet("Xếp-điểm-thi")
    xep_cols = ["Cụm", "Điểm thi"]
    df_xep2 = df_xep_raw[[c for c in xep_cols if c in df_xep_raw.columns]].dropna(subset=["Cụm"])
    df_xep2 = df_xep2[df_xep2["Cụm"].astype(str).str.strip() != ""]
    write_df_sheet(ws_xep, df_xep2)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─── SESSION STATE ─────────────────────────────────────────────────────────────
for key in ["result_df", "result_bytes", "errors", "warnings"]:
    if key not in st.session_state:
        st.session_state[key] = None


# ─── UI ────────────────────────────────────────────────────────────────────────
st.markdown('<div class="app-title">📋 CÁC KỲ THI: TẠO SBD TỰ ĐỘNG</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">Hỗ trợ: BEBRAS · AMC8 · AMC10/12 · VEO</div>', unsafe_allow_html=True)

# ─── SECTION 1: Chọn kỳ thi ────────────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Bước 1</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">🎯 Chọn kỳ thi</div>', unsafe_allow_html=True)

exam_options = {k: v["label"] for k, v in EXAM_RULES.items()}
exam_key = st.radio(
    "Kỳ thi",
    options=list(exam_options.keys()),
    format_func=lambda k: exam_options[k],
    horizontal=True,
    label_visibility="collapsed",
)
st.caption(EXAM_RULES[exam_key]["desc"])
st.markdown('</div>', unsafe_allow_html=True)

# ─── SECTION 2: Upload ──────────────────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Bước 2</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📂 File dữ liệu</div>', unsafe_allow_html=True)

col_dl, col_up = st.columns([1, 2])
with col_dl:
    st.download_button(
        label="⬇️ Tải file mẫu",
        data=load_sample_bytes(),
        file_name="SBD_mẫu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("File mẫu gồm 3 sheets: Học-sinh, Điểm-thi, Xếp-điểm-thi")

with col_up:
    uploaded = st.file_uploader(
        "Upload file data (.xlsx)",
        type=["xlsx"],
        label_visibility="collapsed",
    )

st.markdown('</div>', unsafe_allow_html=True)

# ─── SECTION 3: Xử lý ──────────────────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Bước 3</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">⚙️ Xử lý & Tải kết quả</div>', unsafe_allow_html=True)

if uploaded:
    if st.button("🚀 Tạo SBD ngay", type="primary", use_container_width=False):
        st.session_state.result_df = None
        st.session_state.result_bytes = None
        st.session_state.errors = None

        with st.spinner("Đang xử lý..."):
            try:
                xl = pd.ExcelFile(uploaded)
                if "Học-sinh" not in xl.sheet_names:
                    st.session_state.errors = ["File không có sheet **'Học-sinh'**"]
                elif "Điểm-thi" not in xl.sheet_names:
                    st.session_state.errors = ["File không có sheet **'Điểm-thi'**"]
                elif "Xếp-điểm-thi" not in xl.sheet_names:
                    st.session_state.errors = ["File không có sheet **'Xếp-điểm-thi'**"]
                else:
                    df_hs = pd.read_excel(uploaded, sheet_name="Học-sinh")
                    df_diem = pd.read_excel(uploaded, sheet_name="Điểm-thi")
                    df_xep = pd.read_excel(uploaded, sheet_name="Xếp-điểm-thi")

                    # Drop unnamed helper columns
                    df_hs = df_hs[[c for c in df_hs.columns if not str(c).startswith("Unnamed")]]
                    df_diem = df_diem[[c for c in df_diem.columns if not str(c).startswith("Unnamed")]]
                    df_xep = df_xep[[c for c in df_xep.columns if not str(c).startswith("Unnamed")]]

                    errs = validate_input(df_hs, df_diem, df_xep, exam_key)
                    if errs:
                        st.session_state.errors = errs
                    else:
                        df_result = process_sbd(df_hs, df_diem, df_xep, exam_key)
                        xlsx_bytes = export_excel(df_result, df_diem, df_xep)
                        st.session_state.result_df = df_result
                        st.session_state.result_bytes = xlsx_bytes

            except Exception as e:
                st.session_state.errors = [f"Lỗi không xác định: {str(e)}", f"```\n{traceback.format_exc()}\n```"]

# ── Error display with inline fix ──
if st.session_state.errors:
    for err in st.session_state.errors:
        st.markdown(f'<div class="banner-err">⚠️ {err}</div>', unsafe_allow_html=True)

    st.markdown("#### 🛠️ Khắc phục lỗi ngay trong app")

    # If column errors, might need to remap; show hint
    st.info("Hãy kiểm tra lại file data và chỉnh sửa trực tiếp dưới đây nếu cần:")

    if uploaded and st.session_state.errors:
        try:
            xl2 = pd.ExcelFile(uploaded)
            sheet_fix = st.selectbox("Chọn sheet cần xem/sửa", xl2.sheet_names)
            df_fix = pd.read_excel(uploaded, sheet_name=sheet_fix)
            df_fix = df_fix[[c for c in df_fix.columns if not str(c).startswith("Unnamed")]]

            st.markdown(f"**Nội dung sheet `{sheet_fix}`** (chỉnh sửa trực tiếp):")
            edited = st.data_editor(df_fix, num_rows="dynamic", use_container_width=True, key=f"edit_{sheet_fix}")

            if st.button("🔄 Xử lý lại với dữ liệu đã sửa"):
                try:
                    # We need all 3 sheets; reread others from original
                    all_sheets = {}
                    for s in xl2.sheet_names:
                        all_sheets[s] = pd.read_excel(uploaded, sheet_name=s)
                        all_sheets[s] = all_sheets[s][[c for c in all_sheets[s].columns if not str(c).startswith("Unnamed")]]
                    all_sheets[sheet_fix] = edited

                    df_hs2 = all_sheets.get("Học-sinh", pd.DataFrame())
                    df_diem2 = all_sheets.get("Điểm-thi", pd.DataFrame())
                    df_xep2 = all_sheets.get("Xếp-điểm-thi", pd.DataFrame())

                    errs2 = validate_input(df_hs2, df_diem2, df_xep2, exam_key)
                    if errs2:
                        st.session_state.errors = errs2
                        st.rerun()
                    else:
                        df_result2 = process_sbd(df_hs2, df_diem2, df_xep2, exam_key)
                        xlsx_bytes2 = export_excel(df_result2, df_diem2, df_xep2)
                        st.session_state.result_df = df_result2
                        st.session_state.result_bytes = xlsx_bytes2
                        st.session_state.errors = None
                        st.rerun()
                except Exception as e2:
                    st.error(f"Vẫn còn lỗi: {e2}")
        except Exception:
            pass

# ── Success ──
if st.session_state.result_df is not None:
    df_res = st.session_state.result_df
    st.markdown(f'<div class="banner-ok">✅ Xử lý thành công! Tổng {len(df_res):,} học sinh.</div>', unsafe_allow_html=True)

    # Stats
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng HS", f"{len(df_res):,}")
    if "Điểm thi" in df_res:
        c2.metric("Số điểm thi", df_res["Điểm thi"].nunique())
    if "Phòng thi" in df_res:
        c3.metric("Số phòng thi", df_res.groupby(["Điểm thi","Phòng thi"]).ngroups)
    if "Cấp độ" in df_res:
        c4.metric("Số cấp độ / nhánh", df_res["Cấp độ"].nunique())

    st.dataframe(df_res.head(50), use_container_width=True, height=300)
    if len(df_res) > 50:
        st.caption(f"Hiển thị 50/{len(df_res):,} dòng. Tải file để xem đầy đủ.")

    st.download_button(
        label="📥 Tải file kết quả (.xlsx)",
        data=st.session_state.result_bytes,
        file_name=f"SBD_{exam_key}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=False,
    )

elif uploaded is None:
    st.info("👆 Vui lòng upload file data để bắt đầu.")

st.markdown('</div>', unsafe_allow_html=True)

# ─── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<p style="text-align:center;color:#9CA3AF;font-size:.8rem;">SBD Auto Generator · BEBRAS · AMC8 · AMC10/12 · VEO</p>',
    unsafe_allow_html=True,
)
