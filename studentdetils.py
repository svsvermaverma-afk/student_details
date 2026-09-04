import os
import streamlit as st
import pandas as pd

# पेज सेटअप
st.set_page_config(
    page_title="Class 12 B Student Information",
    page_icon="🎓",
    layout="wide"
)

# डेटा लोड एवं क्लीनिंग
@st.cache_data
def load_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, "XII B INFORMATION.xlsx")
    
    # मुख्य शीट लोड करना
    df = pd.read_excel(file_path, sheet_name="Sheet1 (5)")
    df = df.dropna(subset=["STUDENT'S NAME"]).copy()
    
    # कॉलम नाम साफ़ करना (Line breaks हटाना)
    clean_cols = {}
    for col in df.columns:
        if "OCCUPATION" in str(col):
            clean_cols[col] = "OCCUPATION"
        elif str(col).strip() == "ADDRESS":
            clean_cols[col] = "ADDRESS"
    df.rename(columns=clean_cols, inplace=True)
    
    # 1. नंबर्स और टेक्स्ट फ़ॉर्मेटिंग
    if "ROLL NO." in df.columns:
        df["ROLL NO."] = pd.to_numeric(df["ROLL NO."], errors="coerce").fillna(0).astype(int)
    
    if "S.R. NO." in df.columns:
        df["S.R. NO."] = pd.to_numeric(df["S.R. NO."], errors="coerce").fillna(0).astype(int).astype(str)
        
    if "roll numer 10th" in df.columns:
        df["roll numer 10th"] = df["roll numer 10th"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)

    if "PEN NUMBER" in df.columns:
        df["PEN NUMBER"] = df["PEN NUMBER"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)

    if "AADHAR NO." in df.columns:
        df["AADHAR NO."] = df["AADHAR NO."].fillna("").astype(str)

    if "MOB. NO." in df.columns:
        df["MOB. NO."] = df["MOB. NO."].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)

    if "D.O.B." in df.columns:
        df["D.O.B."] = pd.to_datetime(df["D.O.B."], errors="coerce").dt.strftime("%d-%m-%Y").fillna(df["D.O.B."].astype(str))

    # E.CODE और DEPT की सफ़ाई
    if "E.CODE" in df.columns:
        df["E.CODE"] = df["E.CODE"].fillna("-").astype(str).str.strip().replace("", "-")
    if "DEPT." in df.columns:
        df["DEPT."] = df["DEPT."].fillna("-").astype(str).str.strip().replace("", "-")

    # 2. फ़िल्टर वाले टेक्स्ट कॉलम साफ़ करना
    for col in ["GENDER", "CAT.", "RELIGION", "CASTE", "OCCUPATION"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper().replace("NAN", "-").replace("", "-")

    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"फ़ाइल लोड करने में त्रुटि: {e}")
    st.stop()

# हेडर
st.title("📋 Class XII B - Complete Student Information")
st.caption("Aditya Birla Intermediate College, Renukoot")

# ==================== साइडबार फ़िल्टर्स ====================
st.sidebar.header("🔍 डेटा फ़िल्टर (Data Filters)")

# नाम / पिता का नाम सर्च
search_text = st.sidebar.text_input("विद्यार्थी / पिता का नाम खोजें:")

# ऑक्यूपेशन फ़िल्टर (HE / SUPPLY / OTH)
occ_vals = ["All"] + sorted([x for x in df["OCCUPATION"].unique() if x != "-"])
sel_occ = st.sidebar.selectbox("Occupation (HE / SUPPLY / OTH):", occ_vals)

# जेंडर फ़िल्टर
gender_vals = ["All"] + sorted([x for x in df["GENDER"].unique() if x != "-"])
sel_gender = st.sidebar.selectbox("Gender (लिंग):", gender_vals)

# वर्ग (Category) फ़िल्टर
cat_vals = ["All"] + sorted([x for x in df["CAT."].unique() if x != "-"])
sel_cat = st.sidebar.selectbox("Category (OBC / SC / ST / GEN):", cat_vals)

# धर्म फ़िल्टर
rel_vals = ["All"] + sorted([x for x in df["RELIGION"].unique() if x != "-"])
sel_rel = st.sidebar.selectbox("Religion (धर्म):", rel_vals)

# जाति फ़िल्टर
caste_vals = ["All"] + sorted([x for x in df["CASTE"].unique() if x != "-"])
sel_caste = st.sidebar.selectbox("Caste (जाति):", caste_vals)

# फ़िल्टर लागू करना
filtered_df = df.copy()

if sel_occ != "All":
    filtered_df = filtered_df[filtered_df["OCCUPATION"] == sel_occ]

if sel_gender != "All":
    filtered_df = filtered_df[filtered_df["GENDER"] == sel_gender]

if sel_cat != "All":
    filtered_df = filtered_df[filtered_df["CAT."] == sel_cat]

if sel_rel != "All":
    filtered_df = filtered_df[filtered_df["RELIGION"] == sel_rel]

if sel_caste != "All":
    filtered_df = filtered_df[filtered_df["CASTE"] == sel_caste]

if search_text:
    filtered_df = filtered_df[
        filtered_df["STUDENT'S NAME"].astype(str).str.contains(search_text, case=False, na=False) |
        filtered_df["FATHER'S NAME"].astype(str).str.contains(search_text, case=False, na=False)
    ]

# ==================== शीर्ष समरी कार्ड्स ====================
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("कुल छात्र (Total)", len(filtered_df))
c2.metric("Boys (M)", len(filtered_df[filtered_df["GENDER"] == "M"]))
c3.metric("Girls (F)", len(filtered_df[filtered_df["GENDER"] == "F"]))
c4.metric("Hindalco (HE)", len(filtered_df[filtered_df["OCCUPATION"] == "HE"]))
c5.metric("Supply (HS)", len(filtered_df[filtered_df["OCCUPATION"] == "SUPPLY"]))
c6.metric("Other (OTH)", len(filtered_df[filtered_df["OCCUPATION"] == "OTH"]))

st.markdown("---")

# ==================== इमेज के अनुसार सटीक 20 कॉलम ====================
exact_image_columns = [
    "ROLL NO.", "class", "S.R. NO.", "roll numer 10th", "PEN NUMBER",
    "AADHAR NO.", "D.O.B.", "STUDENT'S NAME", "FATHER'S NAME", "MOTHER'S NAME",
    "GENDER", "CASTE", "CAT.", "RELIGION", "ADDRESS", "MOB. NO.", "EMAIL ID",
    "OCCUPATION", "E.CODE", "DEPT."
]

display_columns = [col for col in exact_image_columns if col in filtered_df.columns]

# मुख्य डेटा टेबल
st.dataframe(
    filtered_df[display_columns],
    use_container_width=True,
    hide_index=True
)

# फ़िल्टर्ड डेटा डाउनलोड बटन
st.download_button(
    label="📥 यह फ़िल्टर किया हुआ डेटा CSV में डाउनलोड करें",
    data=filtered_df[display_columns].to_csv(index=False).encode('utf-8'),
    file_name="Student_Data_Filtered.csv",
    mime="text/csv"
)

# ==================== नीचे सम्पूर्ण सांख्यिकी (Detailed Statistics Breakdown) ====================
st.markdown("---")
st.subheader("📊 वर्गवार एवं ऑक्यूपेशन सांख्यिकी विवरण (Overall Breakdown & Totals)")

stat_col1, stat_col2, stat_col3 = st.columns(3)

with stat_col1:
    st.markdown("##### 📌 सामाजिक वर्ग (Category-wise)")
    cat_summary = filtered_df["CAT."].value_counts().reset_index()
    cat_summary.columns = ["Category", "कुल छात्र (Count)"]
    cat_summary["प्रतिशत (%)"] = ((cat_summary["कुल छात्र (Count)"] / len(filtered_df)) * 100).round(1).astype(str) + "%"
    st.table(cat_summary)

with stat_col2:
    st.markdown("##### 🏢 ऑक्यूपेशन (HE / SUPPLY / OTH)")
    occ_summary = filtered_df["OCCUPATION"].value_counts().reset_index()
    occ_summary.columns = ["Occupation", "कुल छात्र (Count)"]
    occ_summary["विवरण"] = occ_summary["Occupation"].map({
        "HE": "Hindalco Employee (E.Code उपलब्ध)",
        "SUPPLY": "Hindalco Supply",
        "OTH": "Other / Private",
        "-": "उपलब्ध नहीं"
    }).fillna("-")
    st.table(occ_summary)

with stat_col3:
    st.markdown("##### 🚻 लिंग अनुपात (Gender-wise)")
    gen_summary = filtered_df["GENDER"].value_counts().reset_index()
    gen_summary.columns = ["Gender", "कुल छात्र (Count)"]
    gen_summary["प्रतिशत (%)"] = ((gen_summary["कुल छात्र (Count)"] / len(filtered_df)) * 100).round(1).astype(str) + "%"
    st.table(gen_summary)
