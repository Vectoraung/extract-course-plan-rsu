import streamlit as st
from extract_from_rsu36_file.course_plan import CoursePlan
import time

st.title("Course Plan Extractor")

if 'excel' not in st.session_state:
    st.session_state['excel'] = None

if st.session_state['excel'] is None:
    uploaded_file = st.file_uploader("Select your RSU36 file:", type=["pdf"])

    if st.button("Extract"):
        if uploaded_file is not None:
            st.session_state['file'] = uploaded_file
            st.success(f"Course plan extracted successfully! Ready to download.")
            
            cp = CoursePlan()
            cp.import_with_RSU36_file(uploaded_file)
            excel_file, file_name = cp.generate_excel_file(streamlit=True)   # returns BytesIO
            st.session_state['excel'] = (excel_file, file_name)

            st.rerun()
        else:
            empty_container = st.empty()
            empty_container.warning("Please upload a file first.")
            time.sleep(2)
            empty_container.empty()
else:
    st.success(f"Course plan extracted successfully! Ready to download.")

    file, file_name = st.session_state['excel']
    st.download_button(
        label="📥 Download Excel",
        data=file,
        file_name=f"{file_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("Extract Another"):
        st.session_state['excel'] = None

        st.rerun()