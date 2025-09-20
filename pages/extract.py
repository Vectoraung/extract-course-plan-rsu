import streamlit as st
from extract_from_rsu36_file.course_extractor import CourseExtractor
from extract_from_rsu36_file.course_plan_fitter import CoursePlanFitter
import time

st.title("Course Plan Extractor")

st.info("""
**Disclaimer:** This app may produce errors or incorrect results. Please double-check all outputs and use at your own risk.
""")
st.info("""
Currently working only for ICT (Information and Communication Technology) major.
""")

def reset_session_state():
    st.session_state['course_plan_data'] = {
        'excel': None,
        'information': None,
        'ge_courses': None,
        'specialized_courses': None
    }

if 'course_plan_data' not in st.session_state:
    reset_session_state()

if st.session_state['course_plan_data']['excel'] is None:
    
    uploaded_file = st.file_uploader("Select your RSU36 file:", type=["pdf"], )

    if st.button("Extract"):
        if uploaded_file is not None:            
            extractor = CourseExtractor()
            data = extractor.extract_from_rsu36(file=uploaded_file)
            
            fitter = CoursePlanFitter()
            fitter.fit(data)

            excel_file, information, ge_courses, specialized_courses = fitter.generate_excel_file(is_web=True)
            st.session_state['course_plan_data']['excel'] = excel_file
            st.session_state['course_plan_data']['information'] = information
            st.session_state['course_plan_data']['ge_courses'] = ge_courses
            st.session_state['course_plan_data']['specialized_courses'] = specialized_courses
            st.rerun()

        else:
            empty_container = st.empty()
            empty_container.warning("Please upload a file first.")
            time.sleep(2)
            empty_container.empty()
else:
    excel_file = st.session_state['course_plan_data']['excel']
    information = st.session_state['course_plan_data']['information']

    file_name = f"{information['name']}_{information['student_id']}"
    st.success(f"Course plan extracted successfully! ``{file_name}.xlsx`` is ready to download.")

    st.page_link("pages/course_plan_viewer.py", label="View Extracted")

    st.download_button(
        label="Download Excel",
        data=excel_file,
        file_name=f"{file_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("Extract Another"):
        reset_session_state()
        st.rerun()