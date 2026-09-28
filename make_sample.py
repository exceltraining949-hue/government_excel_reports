"""Create a deliberately messy sample government dataset for testing."""
import random
import pandas as pd
import numpy as np

random.seed(42)
np.random.seed(42)

N = 240
depts = ["Finance", "Health", "Education", "Local Government", "Irrigation", "Police"]
dept_variants = {"Finance": ["Finance", "finance", "FINANCE", " Finance "],
                 "Health": ["Health", "health", "HEALTH"],
                 "Education": ["Education", "education", "EDUCATION "],
                 "Local Government": ["Local Government", "local government"],
                 "Irrigation": ["Irrigation", "irrigation"],
                 "Police": ["Police", "police "]}
districts = ["Islamabad", "Rawalpindi", "Lahore", "Karachi", "Peshawar", "Quetta", "Multan", "Faisalabad"]
statuses = ["Completed", "In Progress", "Pending", "Delayed"]
designations = ["Assistant Director", "Deputy Director", "Director", "Section Officer", "Accountant"]
grades = ["BPS-17", "BPS-18", "BPS-19", "BPS-20"]

rows = []
for i in range(N):
    d = random.choice(depts)
    rows.append({
        "Emp_ID": f"EMP-{1000+i}",
        "Employee Name": random.choice(["Ali Raza", "Sana Khan", "Usman Tariq", "Ayesha Malik", "Bilal Ahmed",
                                        "Fatima Noor", "Hamza Sheikh", "Zainab Iqbal", "Kamran Akmal", "Nadia Hussain"]) + f" {i}",
        "Department": random.choice(dept_variants[d]),
        "District": random.choice(districts),
        "Designation": random.choice(designations),
        "Grade": random.choice(grades),
        "Joining Date": pd.Timestamp("2015-01-01") + pd.Timedelta(days=random.randint(0, 3600)),
        "Status": random.choice(statuses),
        "Basic Salary (PKR)": f"Rs. {random.randint(35, 120) * 1000:,}",
        "Budget Allocation": random.randint(2_000_000, 25_000_000),
        "Expenditure": 0,
        "Utilization %": 0,
    })

df = pd.DataFrame(rows)
# expenditure = fraction of budget (a few above budget)
df["Expenditure"] = (df["Budget Allocation"] * np.random.uniform(0.3, 0.95, N)).round(0)
df.loc[df.sample(6, random_state=1).index, "Expenditure"] = df["Budget Allocation"] * 1.15
df["Utilization %"] = (df["Expenditure"] / df["Budget Allocation"] * 100).round(1)
df.loc[10, "Utilization %"] = 130.5   # out-of-range percent

# mess up dates: mixed text formats in a copy column
mixed = df["Joining Date"].copy()
for idx in df.sample(60, random_state=2).index:
    ts = df.loc[idx, "Joining Date"]
    style = idx % 3
    if style == 0:
        mixed.iloc[idx] = ts.strftime("%d/%m/%Y")
    elif style == 1:
        mixed.iloc[idx] = ts.strftime("%d-%b-%Y")
    else:
        mixed.iloc[idx] = ts.strftime("%Y.%m.%d")
df["Joining Date"] = mixed

# inject issues
df.loc[df.sample(10, random_state=3).index, "Department"] = np.nan          # missing department
df.loc[df.sample(8, random_state=4).index, "Employee Name"] = ""            # blank names
df.loc[df.sample(5, random_state=5).index, "Emp_ID"] = np.nan               # missing IDs
df.loc[7, "Emp_ID"] = "EMP-1005"                                           # duplicate ID
df.loc[30, "Emp_ID"] = "EMP-1009"
df.loc[[3, 3.5 if False else 3, 20, 20, 45, 88, 91, 91], :] = df.loc[[3, 3, 20, 20, 45, 88, 91, 91], :].values  # not possible w/ dup index; do below

# exact duplicate rows
dupes = df.sample(12, random_state=6).copy()
df = pd.concat([df, dupes], ignore_index=True)
df = df.sample(frac=1, random_state=7).reset_index(drop=True)

# an error cell and a negative salary string
df.loc[15, "Basic Salary (PKR)"] = "#DIV/0!"
df.loc[25, "Basic Salary (PKR)"] = "Rs. -5,000"
# one extreme outlier expenditure
df.loc[200, "Expenditure"] = 900_000_000
df.loc[200, "Utilization %"] = 4120.0

# second tiny sheet (notes)
notes = pd.DataFrame({"Note": ["This workbook is a test dataset", "Generated 2026-09-28"]})

with pd.ExcelWriter("/home/user/sample_messy_gov_data.xlsx", engine="openpyxl") as w:
    df.to_excel(w, sheet_name="HR_Budget_Data", index=False)
    notes.to_excel(w, sheet_name="Notes", index=False)

print("Sample file created:", len(df), "rows")
print(df.head(3).to_string())
