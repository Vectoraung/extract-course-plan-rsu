import re

class Course:
    def __init__(self, code, credit, grade, term, year_eng, year_thai):
        self.code = code
        self.credit = credit
        self.grade = grade
        self.term = term
        self.year_eng = year_eng
        self.year_thai = year_thai

    def __repr__(self):
        return f"{self.code} ({self.credit} cr) - {self.grade} [{self.term} {self.year_eng}]"


class CourseGroup:
    def __init__(self, name, total_credits, allowed_courses=None, allowed_prefixes=None, allowed_group_numbers=None):
        self.name = name
        self.total_credits = total_credits
        self.allowed_courses = allowed_courses or []  # exact matches
        self.allowed_prefixes = allowed_prefixes or []  # e.g. ["ICT", "IRS"]
        self.allowed_group_numbers = allowed_group_numbers or []  # e.g. [1, 2, 3]
        self.courses = []

    def add_course(self, course):
        code = course.code
        credit = course.credit

        # ✅ Rule 1: exact allowed courses
        if self.allowed_courses and code not in self.allowed_courses:
            return False, f"{code} is not an allowed course for {self.name}"

        # ✅ Rule 2: allowed prefixes
        if self.allowed_prefixes and not any(code.startswith(p) for p in self.allowed_prefixes):
            return False, f"{code} does not match allowed prefixes {self.allowed_prefixes} for {self.name}"

        # ✅ Rule 3: allowed group numbers
        if self.allowed_group_numbers:
            match = re.match(r"^([A-Z]+)(\d+)$", code)
            if match:
                group_number = match.group(2)
                if int(group_number[1]) not in self.allowed_group_numbers:
                    return False, f"{code} has group number {group_number}, not allowed for {self.name}"
            else:
                return False, f"Cannot extract group number from {code}"
                    
            '''try:
                
                # Extract middle digit: e.g. ICT123 -> "123" -> 2
                group_number = int(code[len(code)-3])  
            except (IndexError, ValueError):
                return False, f"Cannot extract group number from {code}"

            if group_number not in self.allowed_group_numbers:
                return False, f"{code} has group number {group_number}, not allowed for {self.name}"'''

        # ✅ Rule 4: credit overflow
        if self.get_total_credits() + credit > self.total_credits:
            return False, f"Adding {code} exceeds credit limit for {self.name}"

        # ✅ Passed all checks → add course
        self.courses.append(course)
        return True, f"{code} successfully added to {self.name}"


    '''def add_course(self, code, credit, grade, term, year_eng, year_thai):
        # ✅ Rule 1: exact allowed courses
        if self.allowed_courses and code not in self.allowed_courses:
            return False, f"{code} is not an allowed course for {self.name}"

        # ✅ Rule 2: allowed prefixes
        if self.allowed_prefixes and not any(code.startswith(p) for p in self.allowed_prefixes):
            return False, f"{code} does not match allowed prefixes {self.allowed_prefixes} for {self.name}"

        # ✅ Rule 3: allowed group numbers
        if self.allowed_group_numbers:
            try:
                # Extract middle digit: e.g. ICT123 -> "123" -> 2
                group_number = int(code[len(code)-3])  
            except (IndexError, ValueError):
                return False, f"Cannot extract group number from {code}"

            if group_number not in self.allowed_group_numbers:
                return False, f"{code} has group number {group_number}, not allowed for {self.name}"

        # ✅ Rule 4: credit overflow
        if self.get_total_credits() + credit > self.total_credits:
            return False, f"Adding {code} exceeds credit limit for {self.name}"

        # ✅ Passed all checks → add course
        self.courses.append(Course(code, credit, grade, term, year_eng, year_thai))
        return True, f"{code} successfully added to {self.name}"'''

    def get_total_credits(self):
        return sum(c.credit for c in self.courses)

    def calculate_gpa(self):
        grade_points = {
            "A": 4.0, "B+": 3.5, "B": 3.0,
            "C+": 2.5, "C": 2.0,
            "D+": 1.5, "D": 1.0, "F": 0.0
        }
        total_points, total_credits = 0, 0
        for c in self.courses:
            if c.grade in grade_points:
                total_points += grade_points[c.grade] * c.credit
                total_credits += c.credit
        return round(total_points / total_credits, 2) if total_credits > 0 else None
    
    def show_courses(self):
        print(f"\n📚 {self.name} (Max Credits: {self.total_credits})")
        print("-" * 60)
        if not self.courses:
            print("No courses added yet.")
            return

        # Sort by code (e.g. ICT101 < ICT205 < IRS111)
        sorted_courses = sorted(self.courses, key=lambda c: c.code)

        for i, c in enumerate(sorted_courses, 1):
            print(f"{i}. {c.code:<8} | {c.credit} Credits | Grade: {c.grade}"
                f" | Term: {c.term} | Year(EN): {c.year_eng} | Year(TH): {c.year_thai}")

        print(f"--> Total Credits so far: {self.get_total_credits()}")

    def convert_to_excel_format(self, show_all_allowed_courses=False):
        """
        Convert courses to a 2D list (Excel-friendly format) with styling:
        - Title row (group name + max credits)
        - Header row
        - Rows for each course
        - Footer row with total credits
        """

        courses = self.courses

        if show_all_allowed_courses:
            # Add missing courses
            finished_courses = [item.code for item in courses]
            for course_code in self.allowed_courses:
                if course_code not in finished_courses:
                    courses.append(
                        Course(
                            code=course_code,
                            credit=0,
                            grade="",
                            term="",
                            year_eng=0,
                            year_thai=0
                        )
                    )

        data = []

        current_total_credits = self.get_total_credits()

        # Title row
        title_text = f"{self.name} (Max Credits: {self.total_credits})"
        note_text = ""
        if current_total_credits == self.total_credits:
            note_text += " (FINISHED!)"
        else:
            note_text += f" (LEFT CREDITS: {self.total_credits - current_total_credits})"
        title = [title_text, "", "", "", "", "", note_text, ""]
        data.append(title)

        # Empty row for spacing
        data.append([])

        # Header row
        header = ["#", "Course Code", "Course Name", "Credits", "Grade", "Term", "Year (EN)", "Year (TH)"]
        data.append(header)

        # Ensure sorted order
        sorted_courses = sorted(courses, key=lambda c: c.code)

        # Course rows
        for i, c in enumerate(sorted_courses, 1):
            row = [
                i,
                c.code,
                getattr(c, "name", ""),  # optional course name if available
                "" if c.credit == 0 else c.credit,
                c.grade,
                c.term,
                "" if c.year_eng == 0 else c.year_eng,
                "" if c.year_thai == 0 else c.year_thai
            ]
            data.append(row)

        # Footer row: total credits
        total_row = ["", "", "Total Credits", current_total_credits]
        data.append([])
        data.append(total_row)

        return data

    def __repr__(self):
        return f"{self.name} ({self.get_total_credits()} / {self.total_credits} credits)"
