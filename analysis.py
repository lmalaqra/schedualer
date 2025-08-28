import pandas as pd
import json
import re

# Load Excel file
df = pd.read_excel("x.xlsx", sheet_name="Sheet1")

# Keywords to look for
keywords = ["عيادة", "عملي", "مختبر"]

# Clinics that are always in the Old Campus
old_campus_clinics = [
    "عيادة طب أسنان الأطفال 1",
    "عيادة استعاضة سنية متحركة 4",
    "عيادة جراحة الفم والأسنان والفكين 1",
    "عيادة طب الأسنان التحفظي 5",
    "عيادة مداواة الأسنان اللبية 4",
    "عيادة علم أمراض اللثة 3"
]

# Clean and normalize session names
def clean_session_name(name):
    if pd.isna(name):
        return ""
    # Replace non-breaking spaces with regular spaces
    name = str(name).replace("\xa0", " ")
    # Remove duplicate repeated substrings (simple heuristic)
    # Split by space, remove duplicates while preserving order
    parts = name.split()
    seen = set()
    clean_parts = []
    for p in parts:
        if p not in seen:
            clean_parts.append(p)
            seen.add(p)
    # Rejoin and strip
    return " ".join(clean_parts).strip()

# Dictionary to hold grouped data
schedule_by_day = {}

# Loop through each row
for _, row in df.iterrows():
    # Combine all row cells into a single string for keyword search
    row_text = " ".join([str(cell) for cell in row if pd.notna(cell)])

    if any(keyword in row_text for keyword in keywords):
        # Detect day
        day_match = re.search(r"(سبت|احد|اثنين|ثلاث|اربعاء|خميس|جمعة)", row_text)
        day = day_match.group(0) if day_match else "غير محدد"

        # Extract time
        time_match = re.search(r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})", row_text)
        start_time, end_time = time_match.groups() if time_match else ("", "")

        # Instructor (last non-null cell)
        instructor = str(row.dropna().iloc[-1])

        # Full session name containing keyword
        session_name = ""
        for cell in row:
            if pd.notna(cell) and any(k in str(cell) for k in keywords):
                session_name = clean_session_name(cell)
                break

        # Determine location
        if any(k in session_name for k in ["عملي", "مختبر"]):
            location = "New Campus"
        elif session_name in old_campus_clinics:
            location = "Old Campus"
        else:
            location = "CELT"

        # Build session entry
        entry = {
            "Clinic": session_name,
            "From": start_time,
            "To": end_time,
            "Instructor": instructor
        }

        # Initialize day if not exists
        if day not in schedule_by_day:
            schedule_by_day[day] = {"New Campus": [], "Old Campus": [], "CELT": []}

        # Append session to the correct location
        schedule_by_day[day][location].append(entry)

# Save to JSON
with open("schedule_by_day_location.json", "w", encoding="utf-8") as f:
    json.dump(schedule_by_day, f, ensure_ascii=False, indent=4)

print("✅ Schedule grouped by day and location saved to schedule_by_day_location.json")