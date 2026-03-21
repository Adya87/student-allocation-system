import pandas as pd
import re

def allocate_coscholastics(file_path, activity_capacity=None, valid_grades=None, output_prefix="activity_allocation"):
    """
    Allocate activities to students based on their preferences
    and predefined activity capacities.

    Parameters:
    ----------
    file_path : str
        Path to the Excel file containing student preference data.

    activity_capacity : dict, optional
        Dictionary defining capacity of each activity per grade.

    valid_grades : list, optional
        List of valid grades. Default is [5, 6, 7, 8].

    output_prefix : str, optional
        Prefix for output files. Default is "activity_allocation".

    Output:
    ------
    Creates separate Excel files for each grade.
    """

    if valid_grades is None:
        valid_grades = [5, 6, 7, 8]

    if activity_capacity is None:
        activity_capacity = {
            "Visual Arts: Art & Craft": {5: 35, 6: 35, 7: 35},
            "Performing Arts: Theatre In Education": {5: 35, 6: 35, 7: 35, 8: 35},
            "Performing Arts: Indian Music": {5: 30, 6: 30, 7: 30, 8: 30},
            "Performing Arts: Indian Dance": {5: 30, 6: 30, 7: 30, 8: 30},
            "Performing Arts: Western Music": {5: 30, 6: 30, 7: 30, 8: 30},
            "Performing Arts: Western Dance": {5: 30, 6: 30, 7: 30, 8: 30},
        }

    def extract_grade_and_section(value):
        match = re.match(r'(\d+)([A-Za-z]+)', str(value).strip())
        if match:
            return int(match.group(1)), match.group(2).upper()
        return None, None

    def extract_grade(value):
        match = re.search(r'\d+', str(value))
        return int(match.group()) if match else None

    def init_remaining_capacities():
        return {
            activity: {grade: cap for grade, cap in caps.items()}
            for activity, caps in activity_capacity.items()
        }

    def allocate(preferences, grade, remaining):
        for pref in preferences:
            if pref in remaining and grade in remaining[pref] and remaining[pref][grade] > 0:
                remaining[pref][grade] -= 1
                return pref
        return "No Allocation"

    # Load file
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    df = df.sort_values(by=df.columns[0])
    is_grade8_format = len(df.columns) == 7

    if is_grade8_format:
        df.columns = ['Timestamp', 'Email', 'Pref1', 'Pref2', 'Pref3', 'Name', 'GradeSection']
        df['Grade'], df['Section'] = zip(*df['GradeSection'].map(extract_grade_and_section))
        df['NameKey'] = df['Name'].str.strip().str.lower()
    else:
        df.columns = [f"Col{i}" for i in range(len(df.columns))]
        df['Name'] = df['Col8']
        df['Grade'] = df['Col9'].map(extract_grade)
        df['Section'] = df['Col10'].astype(str).str.strip()
        df['NameKey'] = df['Name'].str.strip().str.lower()

    df = df.drop_duplicates(subset='NameKey', keep='first')

    # ---------- Term I ----------
    remaining_term1 = init_remaining_capacities()
    allocations_term1 = []

    for _, row in df.iterrows():
        name, grade, section = row['Name'], row['Grade'], row['Section']

        if grade not in valid_grades:
            continue

        prefs = (
            [row['Pref1'], row['Pref2'], row['Pref3']]
            if is_grade8_format or grade == 8
            else [row['Col2'], row['Col3'], row['Col4']]
        )

        prefs = [str(p).strip() for p in prefs if pd.notna(p)]
        alloc1 = allocate(prefs, grade, remaining_term1)
        allocations_term1.append([name, grade, section, alloc1])

    # ---------- Term II ----------
    remaining_term2 = init_remaining_capacities()
    allocations_term2 = []

    for _, row in df.iterrows():
        name, grade, section = row['Name'], row['Grade'], row['Section']

        if grade not in valid_grades:
            continue

        prefs = (
            [row['Pref1'], row['Pref2'], row['Pref3']]
            if is_grade8_format or grade == 8
            else [row['Col5'], row['Col6'], row['Col7']]
        )

        prefs = [str(p).strip() for p in prefs if pd.notna(p)]
        alloc2 = allocate(prefs, grade, remaining_term2)
        allocations_term2.append([name, grade, section, alloc2])

    df_term1 = pd.DataFrame(allocations_term1, columns=["Student Name", "Grade", "Section", "Term I Activity"])
    df_term2 = pd.DataFrame(allocations_term2, columns=["Student Name", "Grade", "Section", "Term II Activity"])

    final_df = pd.merge(df_term1, df_term2, on=["Student Name", "Grade", "Section"])

    for grade in sorted(final_df['Grade'].unique()):
        grade_df = final_df[final_df['Grade'] == grade]
        output_file = f"{output_prefix}_grade_{grade}.xlsx"
        grade_df.to_excel(output_file, index=False)
        print(f"Grade {grade} saved → {output_file}")


if __name__ == "__main__":
    file_path = input("Enter input file path: ")
    allocate_coscholastics(file_path)
