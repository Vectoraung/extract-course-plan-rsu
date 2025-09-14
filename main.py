import streamlit as st
from extract_from_rsu36_file.models import CoursePlan

st.title("File Importer")

uploaded_file = st.file_uploader("Choose a file", type=["pdf"])

if st.button("Process File"):
    if uploaded_file is not None:
        st.success(f"File '{uploaded_file.name}' uploaded successfully!")
        
        cp = CoursePlan()
        cp.import_with_RSU36_file(uploaded_file)
        excel_file = cp.generate_excel_file()   # returns BytesIO

        # Add download button
        st.download_button(
            label="📥 Download Excel",
            data=excel_file,
            file_name="course_plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Please upload a file first.")