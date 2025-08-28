import requests
import pandas as pd
from bs4 import BeautifulSoup
import re
import json

# -------------------------------
# 1️⃣ Fetch the table from website
# -------------------------------
url = "https://zajelbs.najah.edu/servlet/materials"
payload = {"b": 10761}  # adjust as needed

response = requests.post(url, data=payload)
if response.status_code != 200:
    raise Exception(f"POST request failed with status code {response.status_code}")

response.encoding = "windows-1256"
html_content = response.text
soup = BeautifulSoup(html_content, "html.parser")
tables = soup.find_all("table")

# Normalize text
def normalize_text(s):
    if not s:
        return ""
    s = str(s).replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

# Identify table by headers
key_headers = [
    "المساق/ش", "اسم المساق", "س.م", "الأيام", "الساعة",
    "القاعة", "الحرم", "المتطلبات السابقة", "المدرس", "أرقام مساقات مكافئة"
]

target_table = None
for table in tables:
    first_row = table.find("tr")
    if not first_row:
        continue
    cols = [normalize_text(td.get_text()) for td in first_row.find_all("td")]
    if all(any(kh in c for c in cols) for kh in key_headers):
        target_table = table
        break

if target_table is None:
    raise Exception("Could not find the table with the expected headers")

# Extract rows
rows = []
for tr in target_table.find_all("tr")[1:]:
    cols = tr.find_all("td")
    if len(cols) == 0:
        continue
    row = [normalize_text(td.get_text(separator=" ", strip=True)) for td in cols]
    rows.append(row)

# Create DataFrame
df = pd.DataFrame(rows)
if df.shape[1] >= 12:
    df.columns = [
        "Image", "Empty1", "المساق/ش", "اسم المساق", "س.م",
        "الأيام", "الساعة", "القاعة", "الحرم",
        "المتطلبات السابقة", "المدرس", "أرقام مساقات مكافئة"
    ]
    df = df.drop(columns=["Image", "Empty1"])

# -------------------------------
# 2️⃣ Filter only clinics and assign location
# -------------------------------
clinic_keywords = ["عيادة", "عملي", "مختبر"]

old_campus_clinics = [
    "عيادة طب أسنان الأطفال 1",
    "عيادة استعاضة سنية متحركة 4",
    "عيادة جراحة الفم والأسنان والفكين 1",
    "عيادة طب الأسنان التحفظي 5",
    "عيادة مداواة الأسنان اللبية 4",
    "عيادة علم أمراض اللثة 3"
]

def normalize_time(time_str):
    if pd.isna(time_str) or str(time_str).strip() == "":
        return "", ""
    try:
        start, end = time_str.split("-")
        return start.strip(), end.strip()
    except:
        return "", ""

def determine_location(clinic_name):
    if "عملي" in clinic_name or "مختبر" in clinic_name:
        return "New Campus"
    elif clinic_name in old_campus_clinics:
        return "Old Campus"
    else:
        return "CELT"

clinics_schedule = {}

for _, row in df.iterrows():
    row_text = " ".join([str(cell) for cell in row if pd.notna(cell)])
    if not any(k in row_text for k in clinic_keywords):
        continue  # skip non-clinics

    day = row["الأيام"] if pd.notna(row["الأيام"]) else "غير محدد"
    time_str = row["الساعة"] if pd.notna(row["الساعة"]) else ""
    start_time, end_time = normalize_time(time_str)
    room = str(row["القاعة"]).strip() if pd.notna(row["القاعة"]) else ""
    course_name = str(row["اسم المساق"]).strip() if pd.notna(row["اسم المساق"]) else ""
    instructor = str(row["المدرس"]).strip() if pd.notna(row["المدرس"]) else ""

    if not any([course_name, start_time, end_time, room, instructor]):
        continue

    location = determine_location(course_name)

    entry = {
        "Course": course_name,
        "From": start_time,
        "To": end_time,
        "Room": room,
        "Location": location,
        "Instructor": instructor
    }

    if day not in clinics_schedule:
        clinics_schedule[day] = {"New Campus": [], "Old Campus": [], "CELT": []}

    clinics_schedule[day][location].append(entry)

# -------------------------------
# 3️⃣ Worker assignment logic
# -------------------------------
workers = list(range(1, 27))
worker_assignments = {w: [] for w in workers}
worker_day_location = {w: {} for w in workers}

def time_to_minutes(t):
    h, m = map(int, t.split(":"))
    return h*60 + m

def is_overlap(start1, end1, start2, end2):
    return not (end1 <= start2 or end2 <= start1)

assigned_schedule = {}

for day, locations in clinics_schedule.items():
    assigned_schedule[day] = {}
    for location, sessions in locations.items():
        sessions_sorted = sorted(sessions, key=lambda s: time_to_minutes(s["From"]))
        assigned_schedule[day][location] = []

        for session in sessions_sorted:
            if location == "New Campus" and session["Course"] != "طب الأسنان التحفظي 1/ عملي":
                required_workers = 1
            else:
                required_workers = 2

            assigned_workers = []

            # Step 1: Prefer workers already in the same location today
            candidates = [w for w in workers
                          if day in worker_day_location[w] and worker_day_location[w][day] == location]

            # Step 2: Add other workers if needed, sorted by least total shifts
            candidates += sorted([w for w in workers if w not in candidates],
                                 key=lambda w: len(worker_assignments[w]))

            session_start = time_to_minutes(session["From"])
            session_end = time_to_minutes(session["To"])

            for w in candidates:
                if len(assigned_workers) >= required_workers:
                    break
                conflict = any(
                    s_day == day and is_overlap(session_start, session_end, s_start, s_end)
                    for s_day, s_start, s_end in worker_assignments[w]
                )
                if conflict:
                    continue
                worker_assignments[w].append((day, session_start, session_end))
                worker_day_location[w][day] = location
                assigned_workers.append(w)

            while len(assigned_workers) < required_workers:
                assigned_workers.append(None)

            session_copy = session.copy()
            session_copy["Workers"] = assigned_workers
            assigned_schedule[day][location].append(session_copy)

# -------------------------------
# 4️⃣ Save JSON
# -------------------------------
with open("assigned_schedule_updated.json", "w", encoding="utf-8") as f:
    json.dump(assigned_schedule, f, ensure_ascii=False, indent=4)

print("✅ Clinics schedule fetched, workers assigned, and saved to assigned_schedule_updated.json")