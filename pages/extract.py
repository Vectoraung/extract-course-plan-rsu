import streamlit as st
from extract_from_rsu36_file.course_plan import CoursePlan
import time

st.title("Course Plan Extractor")

st.info("""
**Disclaimer:** This app may produce errors or incorrect results. Please double-check all outputs and use at your own risk.
""")

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
            
            try:
                # Attempt to generate the Excel file
                excel_file, file_name = cp.generate_excel_file(streamlit=True)   # returns BytesIO
                st.session_state['excel'] = (excel_file, file_name, cp)
                st.session_state['excel'] = {
                    "file": excel_file,
                    "file_name": file_name,
                    "student_name": cp.student_name,
                    "student_id": cp.student_id,
                    "rsu_i": cp.rsu_i.get_all_courses(),
                    "ic": cp.ic.get_all_courses(),
                    "ge": cp.ge.get_all_courses(),
                    "major": cp.major.get_all_courses(),
                    "major_electives": cp.major_elective.get_all_courses(get_only_finished_courses=True),
                }
                st.rerun()  # refresh the app to show download button
            except Exception as e:
                # Show an error message in the app
                st.error(f"Error generating Excel file: {e}")

        else:
            empty_container = st.empty()
            empty_container.warning("Please upload a file first.")
            time.sleep(2)
            empty_container.empty()
else:
    file = st.session_state['excel']['file']
    file_name = f"{st.session_state['excel']['file_name']}.xlsx"

    st.success(f"Course plan extracted successfully! ``{file_name}`` is ready to download.")

    st.download_button(
        label="Download Excel",
        data=file,
        file_name=f"{file_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("Extract Another"):
        st.session_state['excel'] = None
        st.rerun()