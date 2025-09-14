class Course:
    def __init__(self, raw_line):
        self.code = None
        self.credit = None
        self.grade = None
        self.term = None
        self.year_eng = None
        self.year_thai = None

        if raw_line:
            self._format_from_raw_line(raw_line)

    def _format_from_raw_line(self, line):
        line = line.split(" ")
        self.code = line[1]
        self.credit = int(line[2])

        if len(line) < 4:
            return

        self.grade = line[3]

        sem_and_year = line[4].split("/")

        semester_num, year_thai = int(sem_and_year[0]), int(sem_and_year[1])

        if semester_num == 1:
            self.term = "FIRST SEMESTER"
        elif semester_num == 2:
            self.term = "SECOND SEMESTER"
        elif semester_num == 3:
            self.term = "SUMMER SESSION"
        else: self.term = "unknown"

        self.year_eng = self._thai_year_to_english(thai_year=year_thai)
        self.year_thai = year_thai
    
    def _thai_year_to_english(self, thai_year, month = 1):
        """
        Convert Thai Buddhist Era year to Gregorian year.
        
        Args:
            thai_year (int): Thai year (e.g., 2568)
            month (int): Month number (default=1)
            
        Returns:
            int: English year
        """
        if month <= 3:
            return thai_year - 544
        else:
            return thai_year - 543

    def __repr__(self):
        return f"{self.code} ({self.credit} cr) - {self.grade} [{self.term}]"