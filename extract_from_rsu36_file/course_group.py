class CourseGroup:
    def __init__(self, name, total_credits):
        self.name = name
        self.total_credits = total_credits
        self.courses = []

    def add_course(self, course):
        self.courses.append(course)

    def get_total_credits(self):
        """Sum of credits from all courses in this group"""
        total = 0
        for c in self.courses:
            if c.grade is not None:
                total += c.credit
        return total

    def get_courses_by_term(self, term):
        """Return courses taken in a specific term"""
        return [c for c in self.courses if c.term == term]
    
    def show_courses(self, show_only_finished_courses=False):
        print(f"\n📚 {self.name} (Max Credits: {self.total_credits})")
        print("-" * 60)
        if not self.courses:
            print("No courses added yet.")
            return

        # Sort by code (e.g. ICT101 < ICT205 < IRS111)
        sorted_courses = sorted(self.courses, key=lambda c: c.code)

        # Filter if only finished courses are requested
        if show_only_finished_courses:
            sorted_courses = [c for c in sorted_courses if c.grade is not None]

        if not sorted_courses:
            print("No finished courses yet." if show_only_finished_courses else "No courses added yet.")
            return

        for i, c in enumerate(sorted_courses, 1):
            print(f"{i}. {c.code:<8} | {c.credit} Credits | Grade: {c.grade}"
                f" | Term: {c.term} | Year(EN): {c.year_eng} | Year(TH): {c.year_thai}")

        print(f"--> Total Credits so far: {self.get_total_credits()}")

    def convert_to_excel_format(self, show_only_finished_courses=False):
        """
        Convert courses to a 2D list (Excel-friendly format) with styling:
        - Title row (group name + max credits)
        - Header row
        - Rows for each course
        - Footer row with total credits
        """
        courses = self.courses

        courses = self.courses

        if show_only_finished_courses:
            courses = [c for c in courses if c.grade is not None]

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
                c.credit,
                "" if c.grade is None else c.grade,
                "" if c.term is None else c.term,
                "" if c.year_eng == 0 else c.year_eng,
                "" if c.year_thai == 0 else c.year_thai
            ]
            data.append(row)

        # Footer row: total credits
        total_row = ["", "", "Total Credits", current_total_credits]
        
        data.append(total_row)
        data.append([])

        return data

    def calculate_gpa(self):
        """Weighted GPA for this group"""
        grade_points = {
            "A": 4.0, "B+": 3.5, "B": 3.0,
            "C+": 2.5, "C": 2.0,
            "D+": 1.5, "D": 1.0, "F": 0.0
        }
        total_points = 0
        total_credits = 0
        for c in self.courses:
            if c.grade in grade_points:
                total_points += grade_points[c.grade] * c.credit
                total_credits += c.credit
        return round(total_points / total_credits, 2) if total_credits > 0 else None

    def __repr__(self):
        return f"{self.name} ({self.get_total_credits()} / {self.total_credits} credits)"