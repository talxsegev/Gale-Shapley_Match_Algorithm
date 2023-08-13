import pandas as pd

def gale_shapley(students_pref, companies_pref):
    """Gale-Shapley algorithm for stable matching."""
    # Initialize
    n = len(students_pref)
    free_students = list(students_pref.keys())
    proposals = {student: [] for student in students_pref}
    matched = {}

    while free_students:
        student = free_students.pop()
        student_prefs = students_pref[student]

        for company in student_prefs:
            # Check if student has already proposed to this company
            if company not in proposals[student]:
                proposals[student].append(company)

                # If company is free, accept the proposal
                if company not in matched:
                    matched[company] = student
                    break

                # If company is already matched, check if they prefer the new student
                else:
                    current_match = matched[company]
                    if companies_pref[company].index(student) < companies_pref[company].index(current_match):
                        matched[company] = student
                        free_students.append(current_match)
                        break

    return matched

# Load and review the provided CSV files
students_df = pd.read_csv("Students matching.csv")
companies_df = pd.read_csv("company matching.csv")

student_prefs_dict = {
    row["Full name"]: [col.split('[')[1].split(']')[0] for col in row[1:-1].sort_values().index]
    for _, row in students_df.iterrows()
}

company_prefs_dict = {
    row["Full name"]: [col.split('[')[1].split(']')[0] for col in row[1:-1].sort_values().index]
    for _, row in companies_df.iterrows()
}



matchings = gale_shapley(student_prefs_dict, company_prefs_dict)
print(matchings)