import streamlit as st

# Define pages with custom labels
pg = st.navigation([
    st.Page("pages/extract.py", title="Extract"),
    st.Page("pages/course_plan_viewer.py", title="Extracted Tables"),
    st.Page("pages/about.py", title="About"),
    st.Page("pages/tutorial.py", title="RSU36 File"),
])

# Run the navigation
pg.run()
