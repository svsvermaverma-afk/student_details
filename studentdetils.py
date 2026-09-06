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

# 2. डेटा लोडिंग एवं मर्जिंग फ़ंक्शन
@st.cache_data
def load_all_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # ------------------ A. मुख्य इन्फो शीट लोड करना ------------------
    info_files = [
        os.path.join(base_dir, "XII B INFORMATION_2.xlsx"),
        os.path.join(base_dir, "XII B INFORMATION.xlsx"),
    ] + [f for f in glob.glob(os.path.join(base_dir, "*.xlsx")) if "attandance" not in f.lower() and "attendance" not in f.lower()]
    
    info_path = next((f for f in info_files if os.path.exists(f)), None)
    if not info_path:
        raise FileNotFoundError("मुख्य छात्र विवरण फ़ाइल (XII B INFORMATION.xlsx) नहीं मिली।")

    df_info = pd.read_excel(info_path, sheet_name="Sheet1 (5)")
    df_info = df_info.dropna(subset=["STUDENT'S NAME"]).copy()

    # कॉलम नाम साफ़ करना
    for col in df_info.columns:
        if "OCCUPATION" in str(col):
            df_info.rename(columns={col: "OCCUPATION"}, inplace=True)
        elif str(col).strip() == "ADDRESS":
            df_info.rename(columns={col: "ADDRESS"}, inplace=True)

    # डेटा प्रकार सुधारना
    if "ROLL NO." in df_info.columns:
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

    # ------------------ B. अटेंडेंस शीट लोड करना ------------------
    att_files = [
        os.path.join(base_dir, "attandance.xlsx"),
        os.path.join(base_dir, "attendance.xlsx"),
    ] + [f for f in glob.glob(os.path.join(base_dir, "*att*.xlsx"))]
    
    att_path = next((f for f in att_files if os.path.exists(f)), None)

    att_cols_renamed = []
    latest_pct_col = None

    if att_path:
        df_att = pd.read_excel(att_path, sheet_name=0)
        
        # 'STUDENT NAME' कॉलम खोजना
        name_col = next((c for c in df_att.columns if "STUDENT" in str(c).upper()), None)
        if name_col:
            # क्लीनिंग और कॉलम नाम को स्पष्ट बनाना
            new_cols = {}
            for col in df_att.columns:
                col_str = str(col).strip()
                if "2026-04" in col_str or "Apr" in col_str or "APR" in col_str:
                    new_cols[col] = "APR_ATT"
                elif col_str.upper() == "MAY":
                    new_cols[col] = "MAY_ATT"
                elif col_str.upper() == "JULY":
                    new_cols[col] = "JULY_ATT"
                elif col_str.upper() == "AUG":
                    new_cols[col] = "AUG_ATT"
                elif "PER OUT OF" in col_str.upper() or "%" in col_str or "PERCENT" in col_str.upper():
                    # प्रतिशत वाले कॉलम
                    clean_pname = col_str.replace("PER OUT OF ", "% (").strip() + ")"
                    new_cols[col] = clean_pname
                    latest_pct_col = clean_pname  # आख़िरी प्रतिशत कॉलम को ट्रैक करना
                elif "TOAL" in col_str.upper() or "TOTAL" in col_str.upper():
                    new_cols[col] = f"TOTAL_ATT_{col_str}"
            
            df_att.rename(columns=new_cols, inplace=True)
            
            # मैचिंग के लिए क्लीन नेम बनाना
            df_info["MATCH_NAME"] = df_info["STUDENT'S NAME"].astype(str).str.strip().str.upper()
            df_att["MATCH_NAME"] = df_att[name_col].astype(str).str.strip().str.upper()

            # अटेंडेंस के कॉलम (S NO और नाम छोड़कर)
            att_feature_cols = [c for c in df_att.columns if c not in ["S NO.", "S.NO.", name_col, "MATCH_NAME"]]
            
            # नंबर फ़ॉर्मेटिंग प्रतिशत के लिए
            for p_col in att_feature_cols:
                if "%" in p_col:
                    df_att[p_col] = pd.to_numeric(df_att[p_col], errors="coerce").round(1)

            # दोनों फ़ाइलों को मर्ज करना
            df_merged = pd.merge(df_info, df_att[["MATCH_NAME"] + att_feature_cols], on="MATCH_NAME", how="left")
            df_merged.drop(columns=["MATCH_NAME"], inplace=True)
            return df_merged, att_feature_cols, latest_pct_col

    return df_info, [], None

try:
    df, attendance_cols, latest_pct_col = load_all_data()
except Exception as e:
    st.error(f"डेटा लोड करने में समस्या: {e}")
    st.stop()

# 3. मुख्य हेडर
st.title("🎓 Student Master Information & Attendance Dashboard")
st.caption("Aditya Birla Intermediate College, Renukoot • Real-time Records")

# ==================== साइडबार फ़िल्टर्स ====================
st.sidebar.header("🔍 त्वरित फ़िल्टर (Filters)")

# नाम / पिता का नाम सर्च
search_text = st.sidebar.text_input("छात्र या पिता का नाम खोजें:")

# ऑक्यूपेशन फ़िल्टर
occ_options = ["All"] + sorted([x for x in df["OCCUPATION"].unique() if x != "-"])
sel_occ = st.sidebar.selectbox("Occupation (HE / SUPPLY / OTH):", occ_options)

# जेंडर फ़िल्टर
gender_options = ["All"] + sorted([x for x in df["GENDER"].unique() if x != "-"])
sel_gender = st.sidebar.selectbox("Gender (लिंग):", gender_options)

# केटेगरी फ़िल्टर
cat_options = ["All"] + sorted([x for x in df["CAT."].unique() if x != "-"])
sel_cat = st.sidebar.selectbox("Category (OBC / SC / ST / GEN):", cat_options)

# जाति फ़िल्टर
caste_options = ["All"] + sorted([x for x in df["CASTE"].unique() if x != "-"])
sel_caste = st.sidebar.selectbox("Caste (जाति):", caste_options)

# ==================== अटेंडेंस % फ़िल्टर ====================
st.sidebar.markdown("---")
st.sidebar.header("📊 हाजिरी फ़िल्टर (Attendance %)")

att_filter_mode = "सभी विद्यार्थी"
custom_pct = 75

if latest_pct_col and latest_pct_col in df.columns:
    st.sidebar.write(f"वर्तमान ट्रैकिंग: **{latest_pct_col}**")
    att_filter_mode = st.sidebar.radio(
        "हाजिरी के आधार पर चुनें:",
        ["सभी विद्यार्थी", "75% से कम (< 75% Defaulter)", "50% से कम (< 50% Critical)", "कस्टम प्रतिशत फ़िल्टर"]
    )
    if att_filter_mode == "कस्टम प्रतिशत फ़िल्टर":
        custom_pct = st.sidebar.slider("न्यूनतम प्रतिशत चुनें:", 0, 100, 75)
        att_condition = st.sidebar.selectbox("शर्त:", ["से कम (<)", "से अधिक या बराबर (>=)"])

# फ़िल्टर लागू करना
filtered_df = df.copy()

if sel_occ != "All":
    filtered_df = filtered_df[filtered_df["OCCUPATION"] == sel_occ]
if sel_gender != "All":
    filtered_df = filtered_df[filtered_df["GENDER"] == sel_gender]
if sel_cat != "All":
    filtered_df = filtered_df[filtered_df["CAT."] == sel_cat]
if sel_caste != "All":
    filtered_df = filtered_df[filtered_df["CASTE"] == sel_caste]
if search_text:
    filtered_df = filtered_df[
        filtered_df["STUDENT'S NAME"].astype(str).str.contains(search_text, case=False, na=False) |
        filtered_df["FATHER'S NAME"].astype(str).str.contains(search_text, case=False, na=False)
    ]

# अटेंडेंस फ़िल्टरिंग कंडीशन
if latest_pct_col and latest_pct_col in filtered_df.columns:
    if att_filter_mode == "75% से कम (< 75% Defaulter)":
        filtered_df = filtered_df[filtered_df[latest_pct_col] < 75.0]
    elif att_filter_mode == "50% से कम (< 50% Critical)":
        filtered_df = filtered_df[filtered_df[latest_pct_col] < 50.0]
    elif att_filter_mode == "कस्टम प्रतिशत फ़िल्टर":
        if att_condition == "से कम (<)":
            filtered_df = filtered_df[filtered_df[latest_pct_col] < float(custom_pct)]
        else:
            filtered_df = filtered_df[filtered_df[latest_pct_col] >= float(custom_pct)]

# ==================== शीर्ष समरी कार्ड्स ====================
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("कुल छात्र (Total)", len(filtered_df))
c2.metric("Boys (M)", len(filtered_df[filtered_df["GENDER"] == "M"]))
c3.metric("Girls (F)", len(filtered_df[filtered_df["GENDER"] == "F"]))
c4.metric("Hindalco (HE)", len(filtered_df[filtered_df["OCCUPATION"] == "HE"]))

if latest_pct_col and latest_pct_col in df.columns:
    defaulters_75 = len(filtered_df[filtered_df[latest_pct_col] < 75.0])
    avg_att = round(filtered_df[latest_pct_col].mean(), 1) if len(filtered_df) > 0 else 0
    c5.metric("< 75% डिफ़ॉल्टर", defaulters_75)
    c6.metric("औसत हाजिरी %", f"{avg_att}%")
else:
    c5.metric("Supply (HS)", len(filtered_df[filtered_df["OCCUPATION"] == "SUPPLY"]))
    c6.metric("Other (OTH)", len(filtered_df[filtered_df["OCCUPATION"] == "OTH"]))

st.markdown("---")

# ==================== डेटा टेबल व्यू ====================
base_info_cols = [
    "ROLL NO.", "class", "S.R. NO.", "STUDENT'S NAME", "FATHER'S NAME",
    "GENDER", "CAT.", "CASTE", "MOB. NO.", "OCCUPATION", "E.CODE", "DEPT."
]

# मौजूद इन्फो कॉलम + सभी अटेंडेंस कॉलम
all_available_cols = [c for c in base_info_cols if c in filtered_df.columns] + attendance_cols

# कॉलम चुनने का विकल्प
selected_display_cols = st.multiselect(
    "तालिका में देखने के लिए कॉलम चुनें (कस्टमाइज़ करें):",
    options=all_available_cols,
    default=all_available_cols
)

# टेबल दिखाना
st.dataframe(
    filtered_df[selected_display_cols],
    use_container_width=True,
    hide_index=True
)

# डाउनलोड बटन
st.download_button(
    label="📥 यह फ़िल्टर किया हुआ डेटा (इन्फो + अटेंडेंस) डाउनलोड करें",
    data=filtered_df[selected_display_cols].to_csv(index=False).encode('utf-8'),
    file_name="Student_Attendance_Report.csv",
    mime="text/csv"
)

# ==================== नीचे सांख्यिकी विवरण ====================
st.markdown("---")
st.subheader("📊 सांख्यिकी एवं हाजिरी सारांश (Summary Analytics)")

col_stat1, col_stat2, col_stat3 = st.columns(3)

with col_stat1:
    st.markdown("##### 📌 सामाजिक वर्ग (Category-wise)")
    cat_df = filtered_df["CAT."].value_counts().reset_index()
    cat_df.columns = ["Category", "छात्र संख्या"]
    st.table(cat_df)

with col_stat2:
    st.markdown("##### 🏢 ऑक्यूपेशन (HE / SUPPLY / OTH)")
    occ_df = filtered_df["OCCUPATION"].value_counts().reset_index()
    occ_df.columns = ["Occupation", "छात्र संख्या"]
    st.table(occ_df)

with col_stat3:
    if latest_pct_col and latest_pct_col in filtered_df.columns:
        st.markdown(f"##### 🎯 हाजिरी ब्रैकेट विवरण ({latest_pct_col})")
        above_75 = len(filtered_df[filtered_df[latest_pct_col] >= 75])
        between_50_75 = len(filtered_df[(filtered_df[latest_pct_col] >= 50) & (filtered_df[latest_pct_col] < 75)])
        below_50 = len(filtered_df[filtered_df[latest_pct_col] < 50])
        
        att_summary_df = pd.DataFrame({
            "हाजिरी ब्रैकेट": ["75% या अधिक (सुरक्षित)", "50% से 75% (वार्निंग)", "50% से कम (गंभीर)"],
            "छात्र संख्या": [above_75, between_50_75, below_50]
        })
        st.table(att_summary_df)
    else:
        st.markdown("##### 🚻 लिंग अनुपात (Gender-wise)")
        gen_df = filtered_df["GENDER"].value_counts().reset_index()
        gen_df.columns = ["Gender", "छात्र संख्या"]
        st.table(gen_df)
