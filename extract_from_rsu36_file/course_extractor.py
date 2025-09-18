import pdfplumber
import re
import json

class CourseExtractor:
    def __init__(self):
        self.separate_y = 90

    def extract_from_rsu36(self, file):
        pdf = pdfplumber.open(file)
        page = pdf.pages[0]

        info_sec, courses_sec = self._separate_sections(page)

        extracted_data = {
            "information": None,
            "courses": None
        }

        with open("database/faculty.json", "r", encoding="utf-8") as f:
            data = json.load(f)

        extracted_data['information'] = self._extract_student_info(info_sec, data['faculties'])

        extracted_data['courses'] = self._extract_courses(courses_sec)

        return extracted_data

    def _separate_sections(self, page):
        page_width = page.width
        page_height = page.height

        # Information Section
        top_section = page.crop((0, 0, page_width, self.separate_y))
        information_section = top_section.extract_text().splitlines() if top_section.extract_text() else []

        # Courses Section (Left Half)
        left_section = page.crop((0, self.separate_y, page_width/2, page_height))
        text_left = left_section.extract_text().splitlines() if left_section.extract_text() else []

        # Courses Section (Right Half)
        right_section = page.crop((page_width/2, self.separate_y, page_width, page_height))
        text_right = right_section.extract_text().splitlines() if right_section.extract_text() else []

        courses_section = text_left + text_right

        return information_section, courses_section

    def _extract_student_info(self, information_section, faculties):
        information = {
            "name": None,
            "student_id": None,
            "faculty": None
        }

        combined_text = " ".join(information_section)

        pattern = r"นามสกกล\s+(.*?)\s+รหหสประจจาตหว\s+(\d+)"
        match = re.search(pattern, combined_text, re.DOTALL)  # DOTALL in case there are newlines
        if match:
            information["name"] = match.group(1).strip()
            information["student_id"] = match.group(2).strip()

        for faculty in faculties:
            if faculty['id'] == 0:
                continue

            if faculty['thai_name'] in combined_text:
                information["faculty"] = faculty
                break

        return information
    
    def _extract_courses(self, courses_section):
        courses = []
        for raw_line in courses_section:
            if self._is_course_format(raw_line):
                courses.append(raw_line)

        return courses

    def _is_course_format(self, line):
        parts = line.split(" ")
        if len(parts) < 3:
            return False
        if not parts[0].isnumeric():
            return False
        if not parts[1].isalnum():
            return False
        if not parts[2].isdigit() or not 1 <= int(parts[2]) <= 6:
            return False
        
        return True
    
'''ce = CourseExtractor()
data = ce.extract_from_rsu36(file="6603282.pdf")

print(data)'''