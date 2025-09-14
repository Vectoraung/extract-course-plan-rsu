from extract_from_rsu36_file.course_plan import CoursePlan

cp = CoursePlan()
cp.import_with_RSU36_file("6603282.pdf")

cp.generate_excel_file()