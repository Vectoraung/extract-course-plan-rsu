import streamlit as st

# Define pages with custom labels
pg = st.navigation([
    st.Page("pages/extract.py", title="📥 Extract Course Plan"),
    st.Page("pages/about.py", title="ℹ️ About This App"),
    #st.Page("pages/tutorial.py", title="📖 What's RSU36 File?"),
])

# Run the navigation
pg.run()
