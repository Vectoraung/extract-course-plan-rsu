# Course Plan Creator ( Only for Rangit University, Thailand )
An easy web tool to convert into a course plan table from RSU36 form file from intranet RSU.

## Usage Guide
- Download the RSU36 form file from [RSU's intranet](https://intranet.rsu.ac.th/suWeb/SignIn.aspx).
- Go to this extractor website [Course Plan Extractor →](https://extract-course-plan-rsu.streamlit.app/)
- Go to 'Extract' tab.
- Upload your downloaded RSU36 file then click 'Extract' button.
- You can either download the excel file with course plan tables or go to 'Extracted Tables' tab to view the tables.
- Only avaiable for ICT (Information and Communication Technology) major

## Run Locally
1. Clone the repository
```
git clone https://github.com/Vectoraung/extract-course-plan-rsu.git
cd extract-course-plan-rsu
```

2. Create a virtual environment and activate it (optional but recommended)
```
python -m venv venv
source venv/bin/activate   # On macOS/Linux
venv\Scripts\activate      # On Windows
```

3. Install dependencies
```
pip install -r requirements.txt
```

4. Run the app
```
streamlit run main.py
```