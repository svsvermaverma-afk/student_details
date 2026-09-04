import streamlit as st
import pandas as pd

# पेज का टाइटल और लेआउट सेट करना
st.set_page_config(page_title="Student Information Portal", layout="wide")


@st.cache_data
def load_data():
    file_path = "XII B INFORMATION.xlsx"
    # आपकी मुख्य डेटा शीट लोड करना
    df = pd.read_excel(file_path, sheet_name="Sheet1 (5)")
    df = df.dropna(subset=["STUDENT'S NAME"])

    # कॉलम क्लीनिंग
    df["GENDER"] = df["GENDER"].astype(str).str.strip().str.upper()
    df["CAT."] = df["CAT."].astype(str).str.strip().str.upper()
    df["ROLL NO."] = df["ROLL NO."].fillna(0).astype(int)
    return df


df = load_data()

st.title("🎓 Student Information & Analytics Dashboard")
st.markdown("---")

# 1. साइडबार में फ़िल्टर
st.sidebar.header("🔍 डेटा फ़िल्टर (Data Filter)")

gender_options = ["All"] + sorted(list(df["GENDER"].dropna().unique()))
selected_gender = st.sidebar.selectbox("Gender चुनें:", gender_options)

cat_options = ["All"] + sorted(list(df["CAT."].dropna().unique()))
selected_cat = st.sidebar.selectbox("Category (SC/ST/OBC/GEN) चुनें:", cat_options)

search_name = st.sidebar.text_input("विद्यार्थी का नाम खोजें:")

# फ़िल्टर लागू करना
filtered_df = df.copy()

if selected_gender != "All":
    filtered_df = filtered_df[filtered_df["GENDER"] == selected_gender]

if selected_cat != "All":
    filtered_df = filtered_df[filtered_df["CAT."] == selected_cat]

if search_name:
    filtered_df = filtered_df[filtered_df["STUDENT'S NAME"].str.contains(search_name, case=False, na=False)]

# 2. एक क्लिक में समरी मेट्रिक्स (Summary Cards)
col1, col2, col3, col4, col5 = st.columns(5)

total_students = len(filtered_df)
boys_count = len(filtered_df[filtered_df["GENDER"] == "M"])
girls_count = len(filtered_df[filtered_df["GENDER"] == "F"])
obc_count = len(filtered_df[filtered_df["CAT."] == "OBC"])
sc_count = len(filtered_df[filtered_df["CAT."] == "SC"])
st_count = len(filtered_df[filtered_df["CAT."] == "ST"])
gen_count = len(filtered_df[filtered_df["CAT."] == "GEN"])

col1.metric("कुल विद्यार्थी (Total)", total_students)
col2.metric("Boys (M)", boys_count)
col3.metric("Girls (F)", girls_count)
col4.metric("OBC", obc_count)
col5.metric("SC / ST / GEN", f"{sc_count} / {st_count} / {gen_count}")

st.markdown("---")

# 3. मुख्य डेटा टेबल (Full Data View)
display_cols = [
    "ROLL NO.", "S.R. NO.", "STUDENT'S NAME", "GENDER",
    "CAT.", "CASTE", "FATHER'S NAME", "MOB. NO.", "ADDRESS"
]

available_cols = [c for c in display_cols if c in filtered_df.columns]
st.dataframe(filtered_df[available_cols], use_container_width=True, hide_index=True)

# 4. फ़िल्टर्ड डेटा डाउनलोड करने की सुविधा
csv_data = filtered_df.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 फ़िल्टर किया हुआ डेटा डाउनलोड करें (CSV)",
    data=csv_data,
    file_name='filtered_students.csv',
    mime='text/csv'
)