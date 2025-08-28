import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

# -------------------------------
# 1. Send POST request
# -------------------------------
url = "https://zajelbs.najah.edu/servlet/materials"
payload = {
    "b": 10761  # replace or add more form parameters if needed
}

response = requests.post(url, data=payload)
if response.status_code != 200:
    raise Exception(f"POST request failed with status code {response.status_code}")

# Important: force the correct Arabic encoding
response.encoding = "windows-1256"
html_content = response.text

# -------------------------------
# 2. Parse HTML
# -------------------------------
soup = BeautifulSoup(html_content, "html.parser")
tables = soup.find_all("table")

# -------------------------------
# 3. Normalize text for matching
# -------------------------------
def normalize_text(s):
    if not s:
        return ""
    s = str(s)
    s = s.replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"\s+", " ", s)  # collapse multiple spaces
    return s.strip()

# Key headers to identify the table
key_headers = [
    "المساق/ش", "اسم المساق", "س.م", "الأيام", "الساعة",
    "القاعة", "الحرم", "المتطلبات السابقة", "المدرس", "أرقام مساقات مكافئة"
]

# -------------------------------
# 4. Find the target table by header
# -------------------------------
target_table = None
for i, table in enumerate(tables):
    first_row = table.find("tr")
    if not first_row:
        continue
    cols = [normalize_text(td.get_text()) for td in first_row.find_all("td")]
    # Check if all key headers are present as substrings
    if all(any(kh in c for c in cols) for kh in key_headers):
        target_table = table
        break

if target_table is None:
    # Debug: print headers for all tables
    print("Headers found in tables:")
    for i, table in enumerate(tables):
        first_row = table.find("tr")
        if first_row:
            cols = [normalize_text(td.get_text()) for td in first_row.find_all("td")]
            print(f"Table {i}: {cols}")
    raise Exception("Could not find the table with the expected headers")

print("✅ Target table found!")

# -------------------------------
# 5. Extract rows
# -------------------------------
rows = []
for tr in target_table.find_all("tr")[1:]:  # skip header
    cols = tr.find_all("td")
    if len(cols) == 0:
        continue
    row = [normalize_text(td.get_text(separator=" ", strip=True)) for td in cols]
    rows.append(row)

# -------------------------------
# 6. Convert to DataFrame
# -------------------------------
df = pd.DataFrame(rows)

# Optional: assign column names (skip extra empty columns if present)
if df.shape[1] >= 12:
    df.columns = [
        "Image", "Empty1", "المساق/ش", "اسم المساق", "س.م",
        "الأيام", "الساعة", "القاعة", "الحرم",
        "المتطلبات السابقة", "المدرس", "أرقام مساقات مكافئة"
    ]
    df = df.drop(columns=["Image", "Empty1"])  # remove unused columns

print("✅ Table loaded successfully")
print(df.head())
print(f"Columns: {df.columns.tolist()}")

df.to_excel("t.xlsx")