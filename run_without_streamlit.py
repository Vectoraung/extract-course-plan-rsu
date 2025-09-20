from extract_from_rsu36_file.course_extractor import CourseExtractor
from extract_from_rsu36_file.course_plan_fitter import CoursePlanFitter

ce = CourseExtractor()
data = ce.extract_from_rsu36(file="6603731.pdf")
#print(data)

cp = CoursePlanFitter()
cp.fit(data)

'''dummy_courses = ["ITS191", "FDT102", "IRS155"]
for c in dummy_courses:
    cp.courses.append(
        {
            "code": c,
            "grade": "A",
            "credit": 3,
            "term_number": 1,
            "year_eng": 2025,
            "year_thai": 2568,
            "term": "first / 2025"
        }
    )'''

#for course in cp.courses:
#    print(course)
cp.generate_excel_file()