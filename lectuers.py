import pandas as pd
import json

# Load Excel
df = pd.read_excel("x.xlsx", sheet_name="Sheet1")

# Keywords to exclude (clinic/practical/lab)
exclude_keywords = ["عيادة", "عملي", "مختبر"]

# Result dictionary
other_schedule = {}

def normalize_time(time_str):
    if pd.isna(time_str) or str(time_str).strip() == "":
        return "", ""
    try:
        start, end = time_str.split("-")
        return start.strip(), end.strip()
    except:
        return "", ""

for _, row in df.iterrows():
    row_text = " ".join([str(cell) for cell in row if pd.notna(cell)])
    
    # Skip rows that match clinic/practical/lab keywords
    if any(k in row_text for k in exclude_keywords):
        continue

    # Extract fields with proper NaN handling
    day = row[3] if pd.notna(row[3]) else "غير محدد"
    time_str = row[4] if pd.notna(row[4]) else ""
    start_time, end_time = normalize_time(time_str)
    room = str(row[5]).strip() if pd.notna(row[5]) else ""
    location = str(row[6]).strip() if pd.notna(row[6]) else ""  # "الجديد" or "القديم"
    course_name = str(row[1]).strip() if pd.notna(row[1]) else ""
    instructor = str(row[8]).strip() if pd.notna(row[8]) else ""

    # Skip completely empty rows
    if not any([course_name, start_time, end_time, room, location, instructor]):
        continue

    entry = {
        "Course": course_name,
        "From": start_time,
        "To": end_time,
        "Room": room,
        "Location": location,
        "Instructor": instructor
    }

    # Initialize day
    if day not in other_schedule:
        other_schedule[day] = []

    other_schedule[day].append(entry)

# Save to JSON
with open("other_schedule.json", "w", encoding="utf-8") as f:
    json.dump(other_schedule, f, ensure_ascii=False, indent=4)

print("✅ Other schedule saved to other_schedule.json")