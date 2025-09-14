import streamlit as st

st.title("📖 About This App")

st.markdown("""
Welcome to the **Course Plan Extractor** 🎓  

At **Rangsit University**, students are required to maintain a **custom course plan Excel file**.  
This Excel file helps track which courses have already been completed, so we don’t mistakenly re-add them as if they were new.  

---

### 🎯 Why is this important?
- In the university’s intranet system, you can enroll in courses you’ve already finished — but the system won’t stop you.  
- GPA and grades from all semesters are available online, but **the course groups are not shown**.  
- Course groups matter because each has a **maximum credit limit**.  
  - If you take extra courses in a group that’s already fulfilled, you **don’t get extra credits** — just wasted tuition fees.  
- Manually filling finished courses into Excel is time-consuming and error-prone.  

---

### 🚀 What this app does
This app makes the process **automatic**:  
1. Log in to the RSU intranet and download your `rsu36_form.pdf`.  
2. Upload it here.  
3. Instantly get a **well-structured Excel file** with your completed courses, organized by group.  

No more tedious manual work ✅  
No more worrying about wasted credits ✅  

---

### 👨‍💻 Developer
- **Name:** Aung Khant Kyaw  
- **Major:** ICT, Final Year  
            
📧 **Contact:** [aungkhant.k66@rsu.ac.th](mailto:aungkhantkyaw@gmail.com)  

---

💡 Built with **Python** and **Streamlit** to make student life a little easier.
""")