import os
import glob
import streamlit as st
import pandas as pd

# 1. पेज कॉन्फ़िगरेशन
st.set_page_config(
    page_title="Class 12 B Student Information",
    page_icon="🎓",
    layout="wide"
)

# 2. डेटा लोडिंग फ़ंक्शन (Auto-detecting Excel File)
@st.cache_data
def load_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # फ़ाइल ढूँढने की प्राथमिकता (अगर नाम बदला भी हो तो अपने आप पकड़ लेगा)
    possible_files = [
        os.path.join(base_dir, "XII B INFORMATION_2.xlsx"),
        os.path.join(base_dir, "XII B INFORMATION.xlsx"),
    ] + glob.glob(os.path.join(base_dir, "*.xlsx"))
    
    file_path = None
    for f in possible_files:
        if os.path.exists(f):
            file_path = f
            break
            
    if not file_path:
        raise FileNotFoundError("कोई भी .xlsx एक्सेल फ़ाइल नहीं मिली। कृपया GitHub में फ़ाइल चेक करें।")

    # Sheet1 (5) मुख्य डेटा शीट है
    df = pd.read_excel(file_path, sheet_name="Sheet1 (5)")
    df = df.dropna(subset=["STUDENT'S NAME"]).copy()

    # कॉलम नाम साफ़ करना
    for col in df.columns:
        if "OCCUPATION" in str(col):
            df.rename(columns={col: "OCCUPATION"}, inplace=True)
        elif str(col).strip() == "ADDRESS":
            df.rename(columns={col: "ADDRESS"}, inplace=True)

    # फ़ॉर्मेटिंग
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

    if "E.CODE" in df.columns:
        df["E.CODE"] = df["E.CODE"].fillna("-").astype(str).str.strip().replace("", "-")

    if "DEPT." in df.columns:
        df["DEPT."] = df["DEPT."].fillna("-").astype(str).str.strip().replace("", "-")

    # टेक्स्ट व कैटेगरी कॉलम
    for col in ["GENDER", "CAT.", "RELIGION", "CASTE", "OCCUPATION"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper().replace("NAN", "-").replace("", "-")

    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"एरर: {e}")
    st.stop()

# 3. मुख्य हेडर
st.title("📋 Class XII B - Student Information Dashboard")
st.caption("Aditya Birla Intermediate College, Renukoot")

# 4. साइडबार फ़िल्टर्स
st.sidebar.header("🔍 फ़िल्टर ऑप्शंस (Filters)")

search_text = st.sidebar.text_input("छात्र या पिता का नाम खोजें:")

# ऑक्यूपेशन फ़िल्टर (HE / SUPPLY / OTH)
occ_options = ["All"] + sorted([x for x in df["OCCUPATION"].unique() if x != "-"])
sel_occ = st.sidebar.selectbox("Occupation (HE / SUPPLY / OTH):", occ_options)

# जेंडर फ़िल्टर
gender_options = ["All"] + sorted([x for x in df["GENDER"].unique() if x != "-"])
sel_gender = st.sidebar.selectbox("Gender (लिंग):", gender_options)

# केटेगरी फ़िल्टर
cat_options = ["All"] + sorted([x for x in df["CAT."].unique() if x != "-"])
sel_cat = st.sidebar.selectbox("Category (OBC / SC / ST / GEN):", cat_options)

# धर्म फ़िल्टर
rel_options = ["All"] + sorted([x for x in df["RELIGION"].unique() if x != "-"])
sel_rel = st.sidebar.selectbox("Religion (धर्म):", rel_options)

# जाति फ़िल्टर
caste_options = ["All"] + sorted([x for x in df["CASTE"].unique() if x != "-"])
sel_caste = st.sidebar.selectbox("Caste (जाति):", caste_options)

# फ़िल्टरिंग लागू करना
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

# 5. टॉप समरी कार्ड्स
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("कुल छात्र (Total)", len(filtered_df))
c2.metric("Boys (M)", len(filtered_df[filtered_df["GENDER"] == "M"]))
c3.metric("Girls (F)", len(filtered_df[filtered_df["GENDER"] == "F"]))
c4.metric("Hindalco (HE)", len(filtered_df[filtered_df["OCCUPATION"] == "HE"]))
c5.metric("Supply (HS)", len(filtered_df[filtered_df["OCCUPATION"] == "SUPPLY"]))
c6.metric("Other (OTH)", len(filtered_df[filtered_df["OCCUPATION"] == "OTH"]))

st.markdown("---")

# 6. मुख्य डेटा टेबल (इमेज के 20 कॉलम क्रम में)
exact_image_columns = [
    "ROLL NO.", "class", "S.R. NO.", "roll numer 10th", "PEN NUMBER",
    "AADHAR NO.", "D.O.B.", "STUDENT'S NAME", "FATHER'S NAME", "MOTHER'S NAME",
    "GENDER", "CASTE", "CAT.", "RELIGION", "ADDRESS", "MOB. NO.", "EMAIL ID",
    "OCCUPATION", "E.CODE", "DEPT."
]

display_columns = [c for c in exact_image_columns if c in filtered_df.columns]

st.dataframe(
    filtered_df[display_columns],
    use_container_width=True,
    hide_index=True
)

# डाउनलोड बटन
st.download_button(
    label="📥 फ़िल्टर किया हुआ डेटा CSV में डाउनलोड करें",
    data=filtered_df[display_columns].to_csv(index=False).encode('utf-8'),
    file_name="Filtered_Student_Data.csv",
    mime="text/csv"
)

# 7. नीचे पूरी सांख्यिकी व समरी टेबल्स (Detailed Totals & Statistics)
st.markdown("---")
st.subheader("📊 वर्गवार एवं ऑक्यूपेशन सांख्यिकी (Complete Statistical Summary)")

stat1, stat2, stat3 = st.columns(3)

with stat1:
    st.markdown("##### 📌 सामाजिक वर्ग (Category-wise)")
    cat_df = filtered_df["CAT."].value_counts().reset_index()
    cat_df.columns = ["Category", "कुल छात्र"]
    cat_df["प्रतिशत (%)"] = ((cat_df["कुल छात्र"] / len(filtered_df)) * 100).round(1).astype(str) + "%"
    st.table(cat_df)

with stat2:
    st.markdown("##### 🏢 ऑक्यूपेशन (HE / SUPPLY / OTH)")
    occ_df = filtered_df["OCCUPATION"].value_counts().reset_index()
    occ_df.columns = ["Occupation", "कुल छात्र"]
    occ_df["विवरण"] = occ_df["Occupation"].map({
        "HE": "Hindalco Employee (E.Code व Dept सहित)",
        "SUPPLY": "Hindalco Supply",
        "OTH": "Other / Private",
        "-": "उपलब्ध नहीं"
    }).fillna("-")
    st.table(occ_df)

with stat3:
    st.markdown("##### 🚻 लिंग अनुपात (Gender-wise)")
    gen_df = filtered_df["GENDER"].value_counts().reset_index()
    gen_df.columns = ["Gender", "कुल छात्र"]
    gen_df["प्रतिशत (%)"] = ((gen_df["कुल छात्र"] / len(filtered_df)) * 100).round(1).astype(str) + "%"
    st.table(gen_df)
