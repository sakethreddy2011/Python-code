
import pandas as pd
import re

def detect_absence_streaks(data):
    streak_records = []
    
    for student, records in data.groupby("student_id"):
        absent_days = records.loc[records["status"] == "Absent", "attendance_date"].reset_index(drop=True)
        
        start_date = None
        count = 0
        for i in range(len(absent_days)):
            if i == 0 or (absent_days[i] - absent_days[i - 1]).days == 1:
                if count == 0:
                    start_date = absent_days[i]
                count += 1
            else:
                if count > 3:
                    streak_records.append([student, start_date, absent_days[i - 1], count])
                start_date = absent_days[i]
                count = 1
        
        if count > 3:
            streak_records.append([student, start_date, absent_days.iloc[-1], count])
    
    return pd.DataFrame(streak_records, columns=["student_id", "absence_start_date", "absence_end_date", "total_absent_days"])

def validate_email(email):
    return bool(re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*@[a-zA-Z]+\.com$", str(email)))

def run(path):
    try:
        import pandas as pd
    except ImportError:
        import os
        os.system('pip install pandas openpyxl')
        import pandas as pd
    
    xls = pd.ExcelFile(path)
    attendance_data = pd.read_excel(xls, sheet_name="Attendance_data")
    student_info = pd.read_excel(xls, sheet_name="Student_data")
    
    attendance_data["attendance_date"] = pd.to_datetime(attendance_data["attendance_date"])
    attendance_data.sort_values(by=["student_id", "attendance_date"], inplace=True)
    
    absences = detect_absence_streaks(attendance_data)
    merged_data = absences.merge(student_info, on="student_id", how="left")
    
    merged_data["email"] = merged_data["parent_email"].apply(lambda mail: mail if validate_email(mail) else None)
    
    merged_data["msg"] = merged_data.apply(
        lambda row: (f"Dear Parent, your child {row['student_name']} was absent from "
                     f"{row['absence_start_date'].strftime('%Y-%m-%d')} to {row['absence_end_date'].strftime('%Y-%m-%d')} "
                     f"for {row['total_absent_days']} days. Please ensure their attendance improves.") if row["email"] else None,
        axis=1
    )
    
    return merged_data[["student_id", "absence_start_date", "absence_end_date", "total_absent_days", "email", "msg"]]
