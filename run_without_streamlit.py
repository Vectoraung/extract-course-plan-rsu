from extract_from_rsu36_file.course_extractor import CourseExtractor
from extract_from_rsu36_file.course_plan_fitter import CoursePlanFitter

ce = CourseExtractor()
data = ce.extract_from_rsu36(file="")
#print(data)

cp = CoursePlanFitter()
cp.fit(data)
#for course in cp.courses:
#    print(course)
cp.generate_excel_file()