import os
import glob
import streamlit as st
import pandas as pd

# 1. पेज कॉन्फ़िगरेशन
st.set_page_config(
    page_title="Class 12 B Master Portal",
    page_icon="🎓",
    layout="wide"
)

# 2. डेटा लोडिंग एवं सटीक मर्जिंग फ़ंक्शन
@st.cache_data
def load_all_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    all_excel_files = glob.glob(os.path.join(base_dir, "*.xlsx")) + glob.glob(os.path.join(base_dir, "*.xls"))
    
    # ------------------ A. मुख्य छात्र विवरण शीट ------------------
    info_path = None
    for f in all_excel_files:
        fname = os.path.basename(f).lower()
        if "info" in fname or ("xii" in fname and "att" not in fname and "test" not in fname):
            info_path = f
            break
    if not info_path:
        # फ़ॉलबैक: कोई भी गैर-अटेंडेंस, गैर-टेस्ट फ़ाइल
        for f in all_excel_files:
            fname = os.path.basename(f).lower()
            if "att" not in fname and "test" not in fname:
                info_path = f
                break

    if not info_path:
        raise FileNotFoundError("मुख्य छात्र विवरण फ़ाइल (XII B INFORMATION.xlsx) नहीं मिली।")

    # शीट ढूँढना
    xls_info = pd.ExcelFile(info_path)
    info_sheet = next((s for s in xls_info.sheet_names if "5" in s), xls_info.sheet_names[0])
    df_info = pd.read_excel(info_path, sheet_name=info_sheet)
    df_info = df_info.dropna(subset=["STUDENT'S NAME"]).copy()

    for col in list(df_info.columns):
        if "OCCUPATION" in str(col):
            df_info.rename(columns={col: "OCCUPATION"}, inplace=True)
        elif str(col).strip() == "ADDRESS":
            df_info.rename(columns={col: "ADDRESS"}, inplace=True)

    df_info["ROLL NO."] = pd.to_numeric(df_info["ROLL NO."], errors="coerce").fillna(0).astype(int)
    
    if "S.R. NO." in df_info.columns:
        df_info["S.R. NO."] = pd.to_numeric(df_info["S.R. NO."], errors="coerce").fillna(0).astype(int).astype(str)
    if "roll numer 10th" in df_info.columns:
        df_info["roll numer 10th"] = df_info["roll numer 10th"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    if "PEN NUMBER" in df_info.columns:
        df_info["PEN NUMBER"] = df_info["PEN NUMBER"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    if "AADHAR NO." in df_info.columns:
        df_info["AADHAR NO."] = df_info["AADHAR NO."].fillna("").astype(str)
    if "MOB. NO." in df_info.columns:
        df_info["MOB. NO."] = df_info["MOB. NO."].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    if "D.O.B." in df_info.columns:
        df_info["D.O.B."] = pd.to_datetime(df_info["D.O.B."], errors="coerce").dt.strftime("%d-%m-%Y").fillna(df_info["D.O.B."].astype(str))
    if "E.CODE" in df_info.columns:
        df_info["E.CODE"] = df_info["E.CODE"].fillna("-").astype(str).str.strip().replace("", "-")
    if "DEPT." in df_info.columns:
        df_info["DEPT."] = df_info["DEPT."].fillna("-").astype(str).str.strip().replace("", "-")

    for col in ["GENDER", "CAT.", "RELIGION", "CASTE", "OCCUPATION"]:
        if col in df_info.columns:
            df_info[col] = df_info[col].astype(str).str.strip().str.upper().replace("NAN", "-").replace("", "-")

    df_info = df_info.drop_duplicates(subset=["ROLL NO."]).copy()
    df_info["_KEY_NAME"] = df_info["STUDENT'S NAME"].astype(str).str.replace(".", "", regex=False).str.strip().str.upper()

    # ------------------ B. अटेंडेंस शीट ------------------
    att_path = None
    for f in all_excel_files:
        if "att" in os.path.basename(f).lower():
            att_path = f
            break
            
    attendance_cols = []
    latest_pct_col = None

    if att_path:
        df_att = pd.read_excel(att_path, sheet_name=0)
        sno_col = next((c for c in df_att.columns if "S" in str(c).upper() and "NO" in str(c).upper()), None)
        name_col = next((c for c in df_att.columns if "STUDENT" in str(c).upper()), None)
        
        renamed_att = {}
        for c in df_att.columns:
            c_str = str(c).strip()
            if "2026-04" in c_str or c_str.upper() in ["APR", "APRIL", "APR-"]:
                renamed_att[c] = "APR"
            elif "PER OUT OF" in c_str.upper() or "%" in c_str:
                clean_pct = c_str.replace("PER OUT OF", "% OUT OF")
                renamed_att[c] = clean_pct
                latest_pct_col = clean_pct
            elif "TOAL" in c_str.upper():
                renamed_att[c] = c_str.replace("TOAL", "TOTAL")
            else:
                renamed_att[c] = c_str
        df_att.rename(columns=renamed_att, inplace=True)

        if sno_col and sno_col in df_att.columns:
            df_att["ROLL NO."] = pd.to_numeric(df_att[sno_col], errors="coerce").fillna(0).astype(int)
        else:
            df_att["ROLL NO."] = range(1, len(df_att) + 1)

        df_att = df_att.drop_duplicates(subset=["ROLL NO."]).copy()
        att_cols = [c for c in df_att.columns if c not in ["S NO.", "S.NO.", name_col, "ROLL NO."]]
        for pc in att_cols:
            if "%" in pc:
                df_att[pc] = pd.to_numeric(df_att[pc], errors="coerce").round(1)

        df_info = pd.merge(df_info, df_att[["ROLL NO."] + att_cols], on="ROLL NO.", how="left")
        attendance_cols = att_cols

    # ------------------ C. मंथली टेस्ट शीट (यूनिवर्सल डिटेक्शन) ------------------
    test_path = None
    for f in all_excel_files:
        if "test" in os.path.basename(f).lower():
            test_path = f
            break
            
    test_cols = []
    if test_path:
        xls_test = pd.ExcelFile(test_path)
        # 12वीं की टेस्ट शीट ढूँढना जिसमें HINDI और TOTAL दोनों हों
        best_sheet = None
        for s in xls_test.sheet_names:
            raw_s = pd.read_excel(test_path, sheet_name=s, nrows=6)
            s_text = " ".join([str(v).upper() for v in raw_s.values.flatten()])
            if "HINDI" in s_text and "TOTAL" in s_text:
                best_sheet = s
                break
        if not best_sheet:
            best_sheet = xls_test.sheet_names[0]

        df_raw_test = pd.read_excel(test_path, sheet_name=best_sheet)
        
        # हेडर पंक्ति खोजना
        header_idx = None
        for i in range(min(6, len(df_raw_test))):
            row_str = " ".join([str(x).upper() for x in df_raw_test.iloc[i].values])
            if "HINDI" in row_str and "TOTAL" in row_str:
                header_idx = i
                break

        if header_idx is not None:
            df_test_data = df_raw_test.iloc[header_idx + 1:].copy()
            new_headers = df_raw_test.iloc[header_idx].values.tolist()

            cols_map = {}
            for idx, h in enumerate(new_headers):
                h_str = str(h).strip().upper()
                orig_col = df_test_data.columns[idx]
                if "ROLL" in h_str:
                    cols_map[orig_col] = "ROLL NO."
                elif "HINDI" in h_str:
                    cols_map[orig_col] = "TEST_HINDI (20)"
                elif "ENG" in h_str:
                    cols_map[orig_col] = "TEST_ENG (20)"
                elif "MATH" in h_str:
                    cols_map[orig_col] = "TEST_MATHS (20)"
                elif "PHY" in h_str:
                    cols_map[orig_col] = "TEST_PHY (20)"
                elif "CHE" in h_str:
                    cols_map[orig_col] = "TEST_CHE (20)"
                elif "TOTAL" in h_str:
                    cols_map[orig_col] = "TEST_TOTAL (100)"

            df_test_data.rename(columns=cols_map, inplace=True)

            if "ROLL NO." in df_test_data.columns:
                df_test_data["ROLL NO."] = pd.to_numeric(df_test_data["ROLL NO."], errors="coerce").fillna(0).astype(int)
            else:
                df_test_data["ROLL NO."] = range(1, len(df_test_data) + 1)

            # डुप्लीकेट हटाना (अंतिम रिकॉर्ड रखना)
            df_test_data = df_test_data[df_test_data["ROLL NO."] > 0].drop_duplicates(subset=["ROLL NO."], keep="last").copy()

            subject_cols = [c for c in cols_map.values() if c.startswith("TEST_")]
            for sc in subject_cols:
                df_test_data[sc] = pd.to_numeric(df_test_data[sc], errors="coerce").fillna(0)

            if "TEST_TOTAL (100)" in df_test_data.columns:
                df_test_data["TEST %"] = df_test_data["TEST_TOTAL (100)"].round(1)
                subject_cols.append("TEST %")

            df_info = pd.merge(df_info, df_test_data[["ROLL NO."] + subject_cols], on="ROLL NO.", how="left")
            test_cols = subject_cols

    df_info.drop(columns=["_KEY_NAME"], inplace=True)
    return df_info, attendance_cols, latest_pct_col, test_cols

try:
    df, attendance_cols, latest_pct_col, test_cols = load_all_data()
except Exception as e:
    st.error(f"डेटा लोड करने में समस्या: {e}")
    st.stop()

# 3. मुख्य हेडर
st.title("🎓 Student Master Information, Attendance & Test Portal")
st.caption("Aditya Birla Intermediate College, Renukoot • Real-time Records")

# ==================== साइडबार फ़िल्टर्स ====================
st.sidebar.header("🔍 सामान्य फ़िल्टर (Filters)")

search_text = st.sidebar.text_input("छात्र या पिता का नाम खोजें:")

occ_options = ["All"] + sorted([x for x in df["OCCUPATION"].unique() if x != "-"])
sel_occ = st.sidebar.selectbox("Occupation (HE / SUPPLY / OTH):", occ_options)

gender_options = ["All"] + sorted([x for x in df["GENDER"].unique() if x != "-"])
sel_gender = st.sidebar.selectbox("Gender (लिंग):", gender_options)

cat_options = ["All"] + sorted([x for x in df["CAT."].unique() if x != "-"])
sel_cat = st.sidebar.selectbox("Category (OBC / SC / ST / GEN):", cat_options)

# ==================== अटेंडेंस फ़िल्टर ====================
att_filter_mode = "सभी विद्यार्थी"
custom_att_pct = 75
att_cond = "से कम (<)"

if latest_pct_col and latest_pct_col in df.columns:
    st.sidebar.markdown("---")
    st.sidebar.header("📊 हाजिरी फ़िल्टर (Attendance %)")
    att_filter_mode = st.sidebar.radio(
        f"हाजिरी आधार ({latest_pct_col}):",
        ["सभी विद्यार्थी", "75% से कम (< 75% Defaulter)", "50% से कम (< 50% Critical)", "कस्टम हाजिरी फ़िल्टर"]
    )
    if att_filter_mode == "कस्टम हाजिरी फ़िल्टर":
        custom_att_pct = st.sidebar.slider("न्यूनतम हाजिरी %:", 0, 100, 75)
        att_cond = st.sidebar.selectbox("शर्त (हाजिरी):", ["से कम (<)", "से अधिक या बराबर (>=)"])

# ==================== मंथली टेस्ट फ़िल्टर ====================
test_filter_mode = "सभी विद्यार्थी"
custom_test_pct = 33
test_cond = "से अधिक या बराबर (>=)"

if "TEST %" in df.columns:
    st.sidebar.markdown("---")
    st.sidebar.header("📝 मासिक टेस्ट फ़िल्टर (Test % / Marks)")
    test_filter_mode = st.sidebar.radio(
        "टेस्ट प्रदर्शन आधार पर चुनें:",
        ["सभी विद्यार्थी", "33% से कम (< 33% फेल / कमजोर)", "60% या अधिक (>= 60% First Div)", "75% या अधिक (Distinction)", "कस्टम टेस्ट % फ़िल्टर"]
    )
    if test_filter_mode == "कस्टम टेस्ट % फ़िल्टर":
        custom_test_pct = st.sidebar.slider("टेस्ट प्रतिशत / मार्क्स चुनें:", 0, 100, 40)
        test_cond = st.sidebar.selectbox("शर्त (टेस्ट):", ["से अधिक या बराबर (>=)", "से कम (<)"])
    st.sidebar.success(f"✅ टेस्ट डेटा सक्रिय ({len(test_cols)} फ़ील्ड्स)")
else:
    st.sidebar.warning("⚠️ मंथली टेस्ट फ़ाइल नहीं मिली")

# ==================== फ़िल्टरिंग लागू करना ====================
filtered_df = df.copy()

if sel_occ != "All":
    filtered_df = filtered_df[filtered_df["OCCUPATION"] == sel_occ]
if sel_gender != "All":
    filtered_df = filtered_df[filtered_df["GENDER"] == sel_gender]
if sel_cat != "All":
    filtered_df = filtered_df[filtered_df["CAT."] == sel_cat]
if search_text:
    filtered_df = filtered_df[
        filtered_df["STUDENT'S NAME"].astype(str).str.contains(search_text, case=False, na=False) |
        filtered_df["FATHER'S NAME"].astype(str).str.contains(search_text, case=False, na=False)
    ]

# अटेंडेंस फ़िल्टर
if latest_pct_col and latest_pct_col in filtered_df.columns:
    if att_filter_mode == "75% से कम (< 75% Defaulter)":
        filtered_df = filtered_df[filtered_df[latest_pct_col] < 75.0]
    elif att_filter_mode == "50% से कम (< 50% Critical)":
        filtered_df = filtered_df[filtered_df[latest_pct_col] < 50.0]
    elif att_filter_mode == "कस्टम हाजिरी फ़िल्टर":
        if att_cond == "से कम (<)":
            filtered_df = filtered_df[filtered_df[latest_pct_col] < float(custom_att_pct)]
        else:
            filtered_df = filtered_df[filtered_df[latest_pct_col] >= float(custom_att_pct)]

# टेस्ट फ़िल्टर
if "TEST %" in filtered_df.columns:
    if test_filter_mode == "33% से कम (< 33% फेल / कमजोर)":
        filtered_df = filtered_df[filtered_df["TEST %"] < 33.0]
    elif test_filter_mode == "60% या अधिक (>= 60% First Div)":
        filtered_df = filtered_df[filtered_df["TEST %"] >= 60.0]
    elif test_filter_mode == "75% या अधिक (Distinction)":
        filtered_df = filtered_df[filtered_df["TEST %"] >= 75.0]
    elif test_filter_mode == "कस्टम टेस्ट % फ़िल्टर":
        if test_cond == "से कम (<)":
            filtered_df = filtered_df[filtered_df["TEST %"] < float(custom_test_pct)]
        else:
            filtered_df = filtered_df[filtered_df["TEST %"] >= float(custom_test_pct)]

# ==================== शीर्ष समरी कार्ड्स ====================
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("कुल छात्र (Filtered)", len(filtered_df))
c2.metric("Boys (M)", len(filtered_df[filtered_df["GENDER"] == "M"]))
c3.metric("Girls (F)", len(filtered_df[filtered_df["GENDER"] == "F"]))

if latest_pct_col and latest_pct_col in df.columns:
    defaulters_75 = len(filtered_df[filtered_df[latest_pct_col] < 75.0])
    c4.metric("< 75% हाजिरी", defaulters_75)
else:
    c4.metric("Hindalco (HE)", len(filtered_df[filtered_df["OCCUPATION"] == "HE"]))

if "TEST %" in df.columns:
    pass_count = len(filtered_df[filtered_df["TEST %"] >= 33.0])
    avg_test = round(filtered_df["TEST %"].mean(), 1) if len(filtered_df) > 0 else 0
    c5.metric("टेस्ट पास (>=33%)", pass_count)
    c6.metric("औसत टेस्ट %", f"{avg_test}%")
else:
    c5.metric("Supply (HS)", len(filtered_df[filtered_df["OCCUPATION"] == "SUPPLY"]))
    c6.metric("Other (OTH)", len(filtered_df[filtered_df["OCCUPATION"] == "OTH"]))

st.markdown("---")

# ==================== 3 अलग-अलग टैब (Tabs View) ====================
tab1, tab2, tab3 = st.tabs([
    "📘 1. समग्र मास्टर टेबल (All Data)",
    "📝 2. मंथली टेस्ट मार्क्स (Monthly Test)",
    "📅 3. हाजिरी रिकॉर्ड (Attendance)"
])

# टैब 1: All Data
with tab1:
    base_info_cols = [
        "ROLL NO.", "class", "S.R. NO.", "STUDENT'S NAME", "FATHER'S NAME",
        "GENDER", "CAT.", "CASTE", "MOB. NO.", "OCCUPATION", "E.CODE", "DEPT."
    ]
    available_base = [c for c in base_info_cols if c in filtered_df.columns]
    all_display_options = available_base + attendance_cols + test_cols

    selected_display_cols = st.multiselect(
        "तालिका में कॉलम चुनें:",
        options=all_display_options,
        default=all_display_options,
        key="master_cols"
    )
    st.dataframe(filtered_df[selected_display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)

# टैब 2: Monthly Test Dedicated View
with tab2:
    if test_cols:
        test_view_cols = ["ROLL NO.", "STUDENT'S NAME", "FATHER'S NAME"] + test_cols
        test_view_cols = [c for c in test_view_cols if c in filtered_df.columns]
        st.dataframe(filtered_df[test_view_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
    else:
        st.info("मंथली टेस्ट का डेटा लोड नहीं है।")

# टैब 3: Attendance Dedicated View
with tab3:
    if attendance_cols:
        att_view_cols = ["ROLL NO.", "STUDENT'S NAME", "FATHER'S NAME"] + attendance_cols
        att_view_cols = [c for c in att_view_cols if c in filtered_df.columns]
        st.dataframe(filtered_df[att_view_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
    else:
        st.info("अटेंडेंस डेटा लोड नहीं है।")

# डाउनलोड बटन
st.download_button(
    label="📥 यह फ़िल्टर किया हुआ डेटा (CSV) डाउनलोड करें",
    data=filtered_df.to_csv(index=False).encode('utf-8'),
    file_name="Student_Complete_Master_Report.csv",
    mime="text/csv"
)

# ==================== सांख्यिकी एवं विश्लेषण ====================
st.markdown("---")
st.subheader("📊 समग्र सांख्यिकी एवं प्रदर्शन सारांश (Complete Analytics)")

stat1, stat2, stat3 = st.columns(3)

with stat1:
    st.markdown("##### 📌 सामाजिक वर्ग (Category-wise)")
    cat_df = filtered_df["CAT."].value_counts().reset_index()
    cat_df.columns = ["Category", "छात्र संख्या"]
    st.table(cat_df)

with stat2:
    if latest_pct_col and latest_pct_col in filtered_df.columns:
        st.markdown(f"##### 🎯 हाजिरी ब्रैकेट ({latest_pct_col})")
        above_75 = len(filtered_df[filtered_df[latest_pct_col] >= 75])
        between_50_75 = len(filtered_df[(filtered_df[latest_pct_col] >= 50) & (filtered_df[latest_pct_col] < 75)])
        below_50 = len(filtered_df[filtered_df[latest_pct_col] < 50])
        
        att_summary_df = pd.DataFrame({
            "हाजिरी ब्रैकेट": [">= 75% (सुरक्षित)", "50%–75% (वार्निंग)", "< 50% (गंभीर)"],
            "छात्र संख्या": [above_75, between_50_75, below_50]
        })
        st.table(att_summary_df)

with stat3:
    if "TEST %" in filtered_df.columns:
        st.markdown("##### 📝 मासिक टेस्ट ग्रेड ब्रैकेट (Test %)")
        t_above_75 = len(filtered_df[filtered_df["TEST %"] >= 75])
        t_60_75 = len(filtered_df[(filtered_df["TEST %"] >= 60) & (filtered_df["TEST %"] < 75)])
        t_33_60 = len(filtered_df[(filtered_df["TEST %"] >= 33) & (filtered_df["TEST %"] < 60)])
        t_below_33 = len(filtered_df[filtered_df["TEST %"] < 33])
        
        test_summary_df = pd.DataFrame({
            "टेस्ट ब्रैकेट": [">= 75% (Distinction)", "60%–74% (1st Div)", "33%–59% (Pass)", "< 33% (फेल / कमजोर)"],
            "छात्र संख्या": [t_above_75, t_60_75, t_33_60, t_below_33]
        })
        st.table(test_summary_df)
