import streamlit as st

st.title("How to Download Your RSU36 Form")

st.markdown("""
Follow these steps to download your `rsu36_form.pdf` from the university intranet.
""")

# Step 1
st.header("Step 1: Select Student Login")
st.write("Select the student login option.")
st.image("media/images/step1.png", use_container_width=True)  # replace with your image path

# Step 2
st.header("Step 2: Login")
st.write("Login with your student ID (u6xxxxxx) and password. Make sure you've selected Thai language version (TH).")
st.image("media/images/step2.png", use_container_width=True)

# Step 3
st.header("Step 3: Select Grades By Curriculum Structure")
st.write("Select Grade Information > Grades By Curriculum Structure.")
st.image("media/images/step3.png", use_container_width=True)

# Step 4
st.header("Step 4: Print")
st.write("Click the 'Print' button written in Thai.")
st.image("media/images/step4.png", use_container_width=True)

# Step 5
st.header("Step 5: Download")
st.write("Download the file by clicking the down-arrow button at the top right cornor.")
st.image("media/images/step5.png", use_container_width=True)

st.markdown("---")
st.info("💡 Tip: Make sure the PDF is the official RSU36 form downloaded from the intranet.")
