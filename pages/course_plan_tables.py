import streamlit as st
import pandas as pd

st.title("Extracted Course Plan")

if 'excel' not in st.session_state:
    st.session_state['excel'] = None

if st.session_state['excel'] is None:
    st.warning("Please extract a course plan first.")
    st.stop()

cp = st.session_state['excel'][2]

student_name = cp.student_name
student_id = cp.student_id

# Display student info above the table
st.header(f"Student: {student_name} ({student_id})")

# Prepare header row
header = ["Code", "Credit", "Grade", "Semester", "Year"]

rsu_is = cp.rsu_i.get_all_courses()
table_data = [[
    course.code, course.credit, course.grade, course.term, course.year_eng
] for course in rsu_is]
table_data.insert(0, header)
st.subheader("RSU Identities Courses")
st.table(table_data)

ics = cp.ic.get_all_courses()
table_data = [[
    course.code, course.credit, course.grade, course.term, course.year_eng
] for course in ics]
table_data.insert(0, header)
st.subheader("Internationalization & Communication Courses")
st.table(table_data)

ges = cp.ge.get_all_courses()
table_data = [[
    course.code, course.credit, course.grade, course.term, course.year_eng
] for course in ges]
table_data.insert(0, header)
st.subheader("General Education Courses")
st.table(table_data)

majors = cp.major.get_all_courses()
table_data = [[
    course.code, course.credit, course.grade, course.term, course.year_eng
] for course in majors]
table_data.insert(0, header)
st.subheader("Major Courses")
st.table(table_data)

major_electives = cp.major_elective.get_all_courses(get_only_finished_courses=True)
table_data = [[
    course.code, course.credit, course.grade, course.term, course.year_eng
] for course in major_electives]
table_data.insert(0, header)
st.subheader("Major Elective Courses")
st.table(table_data)