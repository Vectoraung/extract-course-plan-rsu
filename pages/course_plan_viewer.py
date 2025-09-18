import streamlit as st
import pandas as pd

def courses_to_dict(courses: list[dict]) -> dict:
    """
    Convert list of course dicts into a dict of lists,
    using column mapping (with names and order).
    """

    col_map = {
        "code": {"col": 2, "name": "Code"},
        "credit": {"col": 3, "name": "Credit"},
        "grade": {"col": 4, "name": "Grade"},
        "term": {"col": 5, "name": "Term"},
    }

    result = {}

    # Respect order defined in col_map["col"]
    for key in sorted(col_map.keys(), key=lambda k: col_map[k]["col"]):
        col_name = col_map[key]["name"]
        result[col_name] = [str(c.get(key, None)) for c in courses]

    return result

st.title("Course Plan Viewer")

if not 'course_plan_data' in st.session_state:
    st.info("Please extract a course plan first.")
    st.page_link("pages/extract.py", label="Go to Extract")
    st.stop()

if st.session_state['course_plan_data']['ge_courses'] is None or \
    st.session_state['course_plan_data']['specialized_courses'] is None or \
    st.session_state['course_plan_data']['information'] is None:
    
    st.info("Please extract a course plan first.")
    st.page_link("pages/extract.py", label="Go to Extract")
    st.stop()

ge_courses = st.session_state['course_plan_data']['ge_courses']
specialized_courses = st.session_state['course_plan_data']['specialized_courses']
information = st.session_state['course_plan_data']['information']

st.header("General Education Courses")

if ge_courses is None:
    st.warning("No general education courses found.")

for group_name, courses in ge_courses.items():
    st.subheader(group_name)

    courses_dict = courses_to_dict(courses)
    df = pd.DataFrame(courses_dict)
    st.table(df)



st.header(f"Specialized Courses ({information['faculty']['name']})")

if specialized_courses is None:
    st.warning("No specialized courses found.")

for group_name, courses in specialized_courses.items():
    st.subheader(group_name)

    courses_dict = courses_to_dict(courses)
    df = pd.DataFrame(courses_dict)
    st.table(df)