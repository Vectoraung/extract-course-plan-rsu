import pdfplumber
import re

def format_subject(subject_row, semester_name, year_eng, year_thai):
    subject_row = [x for x in subject_row if x is not None]
    return {
        "code": subject_row[0],
        "name": subject_row[1],
        "credit": int(subject_row[2]),
        "grade": subject_row[3],
        "semester": semester_name,
        "year_eng": year_eng,
        "year_thai": year_thai
        }

def format_semester(semester_row):
    semester_row = semester_row[0].split(" ")
    return {
        "semester": semester_row[0],
        "year_eng": semester_row[2],
        "year_thai": semester_row[-1]
}

def clean_row(row):
    """Remove all None or empty string values."""
    return [cell for cell in row if cell not in (None, '')]

def is_subject_row(row):
    """Check if a row starts with a subject code pattern XXX000."""
    if not row:
        return False
    return bool(re.match(r'^[A-Z]{2,3}\d{3}', row[0]))

def is_special_row(row):
    """Detect Semester, Cumulative, or STATUS rows."""
    if not row:
        return False
    first_cell = row[0]
    return first_cell in ("Semester", "Cumulative", "STATUS")

def start_scrapping(pdf_name):
    merged_tables = []
    current_table = []

    with pdfplumber.open(pdf_name) as pdf:
        pages = pdf.pages
        for page in pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    row_clean = clean_row(row)
                    if not row_clean:
                        continue

                    first_cell = row_clean[0]
                    is_new_semester = first_cell and ("SEMESTER" in first_cell or "SESSION" in first_cell)

                    if is_new_semester:
                        if current_table:
                            merged_tables.append(current_table)
                        current_table = [row_clean]  # start new table
                    else:
                        current_table.append(row_clean)  # continuation

    if current_table:
        merged_tables.append(current_table)


    #Combine split subject rows while avoiding special rows
    final_tables = []

    for table in merged_tables:
        new_table = []
        buffer_row = []
        for row in table:
            if is_subject_row(row):
                if buffer_row:
                    new_table.append(buffer_row)
                buffer_row = row
            elif is_special_row(row):
                if buffer_row:
                    new_table.append(buffer_row)
                    buffer_row = []
                new_table.append(row)  # add special row as-is
            else:
                # continuation of previous subject
                if buffer_row:
                    buffer_row += row
                else:
                    # In case a row appears without a subject, just keep it
                    new_table.append(row)
        if buffer_row:
            new_table.append(buffer_row)
        final_tables.append(new_table)

    return final_tables

def get_all_subjects(tables):
    subjects = []
    for table in tables:
        semester_info_text = table[0][0].split(" ")
        semester_name = semester_info_text[0]
        year_eng = semester_info_text[2]
        year_thai = semester_info_text[-1]
        for row in table:
            if is_subject_row(row):
                subjects.append(
                    format_subject(row, semester_name, year_eng, year_thai)
                )
    return subjects

'''final_tables = start_scrapping("Rangsit University.pdf")
final_tables = final_tables[1:]
all_subjects = get_all_subjects(final_tables)

for subject in all_subjects:
    print(subject)

print(len(all_subjects))'''