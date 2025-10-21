from datetime import datetime, timedelta
import pandas as pd
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill

# -------------------------------
# 1️⃣ Define people, work schedules, groups
# -------------------------------
people = [
    "Alice","Bob","Charlie","Diana","Ethan","Fiona",
    "George","Hannah","Ian","Jane","Karl","Liam",
    "Mona","Nina","Oliver","Paula"
]

# Work schedules (0=Mon, ..., 4=Fri)
work_schedule = {
    "Alice": [0,1,2,3],
    "Bob": [0,1,2,3,4],
    "Charlie": [0,1,2,3,4],
    "Diana": [0,1,2,3,4],
    "Ethan": [0,1,2,3,4],
    "Fiona": [0,1,2,3],
    "George": [0,1,2,3,4],
    "Hannah": [0,1,2],
    "Ian": [0,1,2,3,4],
    "Jane": [0,1,2,3,4],
    "Karl": [0,1,2,3,4],
    "Liam": [0,1,2,3,4],
    "Mona": [0,1,2,3,4],
    "Nina": [0,1,2,3,4],
    "Oliver": [0,1,2,3,4],
    "Paula": [0,1,2,3,4],
}

# Groups
southmead_group = ["Alice","Bob","Charlie","Diana","Ethan","Fiona","Karl","Liam"]
uhbw_group = ["George","Hannah","Ian","Jane","Mona","Nina","Oliver","Paula"]

# Cannot swap weekend site - choose anyone who can't work a the other site on the weekend.
cannot_swap_weekend_site = ["Alice","Charlie","George","Mona"]

# Example unavailable dates
unavailable = {
    "Alice": {"2026-03-10":"Annual Leave","2026-05-15":"Conference"},
    "Bob": {"2026-04-22":"Training"},
    "Charlie": {},
    "Diana": {"2026-06-03":"Annual Leave","2026-06-04":"Annual Leave"},
    "Ethan": {},
    "Fiona": {"2026-07-12":"Personal Leave"},
    "George": {},
    "Hannah": {"2026-02-25":"Sick Leave","2026-02-26":"Sick Leave"},
    "Ian": {},
    "Jane": {"2026-04-08":"Conference"},
    "Karl": {},
    "Liam": {},
    "Mona": {},
    "Nina": {},
    "Oliver": {},
    "Paula": {},
}

# -------------------------------
# 2️⃣ Define rota dates
# -------------------------------
start_date = datetime(2026, 2, 4)
end_date = datetime(2026, 8, 5)

# Validate unavailable dates
for person, dates in unavailable.items():
    for ud in dates:
        try:
            date_obj = datetime.strptime(ud, "%Y-%m-%d")
            if not (start_date <= date_obj <= end_date):
                print(f"⚠️ Warning: {person} unavailable date {ud} is outside the rota period.")
        except ValueError:
            print(f"❌ Invalid date format for {person}: {ud}")

# List of all dates
all_dates = []
current_day = start_date
while current_day <= end_date:
    all_dates.append(current_day)
    current_day += timedelta(days=1)

# -------------------------------
# 3️⃣ FTE calculation
# -------------------------------
fte = {p: len(work_schedule[p])/5 for p in people}
total_fte = sum(fte.values())

# -------------------------------
# 4️⃣ Weekend allocation
# -------------------------------
weekend_blocks = [(d, d + timedelta(days=1)) for d in all_dates if d.weekday()==5 and (d+timedelta(days=1))<=end_date]

#Multiply by 2 as two weekend sites to be covered
total_weekends = len(weekend_blocks)*2

# Target weekends (FTE weighted)
target_weekends = {p: total_weekends*(fte[p]/total_fte) for p in people}

# Allocate weekends
weekend_assigned_southmead = defaultdict(int)
weekend_assigned_uhbw = defaultdict(int)
weekend_rota = {}

for sat, sun in weekend_blocks:
    # Southmead
    available_sm = [p for p in southmead_group
                    if all(d.strftime("%Y-%m-%d") not in unavailable.get(p,{}) for d in [sat,sun])]
    # UHBW
    available_uh = [p for p in uhbw_group
                    if all(d.strftime("%Y-%m-%d") not in unavailable.get(p,{}) for d in [sat,sun])]
    # Swap if needed
    if not available_sm:
        available_sm = [p for p in uhbw_group if p not in cannot_swap_weekend_site
                        and all(d.strftime("%Y-%m-%d") not in unavailable.get(p,{}) for d in [sat,sun])]
    if not available_uh:
        available_uh = [p for p in southmead_group if p not in cannot_swap_weekend_site
                        and all(d.strftime("%Y-%m-%d") not in unavailable.get(p,{}) for d in [sat,sun])]

    # Select by fewest weekends / FTE
    chosen_sm = min(available_sm, key=lambda x: weekend_assigned_southmead[x]/fte[x])
    chosen_uh = min(available_uh, key=lambda x: weekend_assigned_uhbw[x]/fte[x])

    weekend_rota[sat.strftime("%Y-%m-%d")] = {"Southmead": chosen_sm, "UHBW": chosen_uh}
    weekend_rota[sun.strftime("%Y-%m-%d")] = {"Southmead": chosen_sm, "UHBW": chosen_uh}

    weekend_assigned_southmead[chosen_sm] += 1
    weekend_assigned_uhbw[chosen_uh] += 1

# Weekend days count for weekday balancing
assigned_counts = {p: 2*weekend_assigned_southmead.get(p,0) + 2*weekend_assigned_uhbw.get(p,0) for p in people}

#------------------------------
# 5.5 Build some protection in for the weekdays before/after a weekend shift
#-----------------------------

weekend_protection = defaultdict(set)

for sat, sun in weekend_blocks:
    weekend_dates = [sat, sun]
    # Restricted weekdays before the weekend
    for i in range(1, 5):  # 1-4 days before
        d_before = sat - timedelta(days=i)
        if d_before.weekday()<5 and start_date <= d_before <= end_date:
            weekend_dates.append(d_before)
    # Restricted weekdays after the weekend
    for i in range(1,6):  # 1-5 days after
        d_after = sun + timedelta(days=i)
        if d_after.weekday()<5 and start_date <= d_after <= end_date:
            weekend_dates.append(d_after)
    # Add to the protection set for both people on that weekend
    for p in [weekend_rota[sat.strftime("%Y-%m-%d")]["Southmead"],
              weekend_rota[sat.strftime("%Y-%m-%d")]["UHBW"]]:
        weekend_protection[p].update(d.strftime("%Y-%m-%d") for d in weekend_dates)



# -------------------------------
# 5️⃣ Weekday allocation
# -------------------------------
total_weekdays = sum(1 for d in all_dates if d.weekday()<5)
full_time_target = total_weekdays / total_fte
target_shifts = {p: full_time_target*fte[p] for p in people}

# Track assigned weekdays (starting from weekend load)
assigned_counts = defaultdict(int, assigned_counts)
rota = {}

for d in all_dates:
    if d.weekday()>=5:
        continue
    date_str = d.strftime("%Y-%m-%d")

    available_people = [
        p for p in people
        if (
            d.weekday() in work_schedule[p] and
            assigned_counts[p] < target_shifts[p] and
            date_str not in unavailable.get(p,{}) and
            date_str not in weekend_protection.get(p,set())
        )
    ]

    if not available_people:
        # fallback: anyone who can work, ignoring FTE
        working_people = [
            p for p in people
            if d.weekday() in work_schedule[p] and
               date_str not in unavailable.get(p,{}) and
               date_str not in weekend_protection.get(p,set())
        ]
        available_people = sorted(working_people, key=lambda x: assigned_counts[x]/fte[x])

    chosen = min(available_people, key=lambda x: assigned_counts[x])
    rota[date_str] = chosen
    assigned_counts[chosen] += 1

# -------------------------------
# 6️⃣ Prepare DataFrame
# -------------------------------
dates_str = [d.strftime("%Y-%m-%d") for d in all_dates]
weekday_names = [d.strftime("%A") for d in all_dates]
data = {"Date": dates_str, "Day": weekday_names}
for p in people:
    data[p] = []

for d in all_dates:
    date_str = d.strftime("%Y-%m-%d")
    day_assignment = rota.get(date_str) or weekend_rota.get(date_str)
    for person in people:
        if date_str in unavailable.get(person,{}):
            data[person].append(unavailable[person][date_str])
        elif isinstance(day_assignment, dict):  # weekend
            if person == day_assignment.get("Southmead"):
                data[person].append("WC SM")
            elif person == day_assignment.get("UHBW"):
                data[person].append("WC UHB")
            else:
                data[person].append("")
        else:
            if person == day_assignment:
                data[person].append("On Call")
            else:
                data[person].append("")

df = pd.DataFrame(data)

excel_filename = "on_call_rota_16people_colored.xlsx"
df.to_excel(excel_filename, index=False)

import openpyxl
from openpyxl.styles import PatternFill

# Open workbook
wb = openpyxl.load_workbook(excel_filename)
ws = wb.active

# Define colors
color_map = {
    "On Call": "90EE90",      # Green
    "WC SM": "ADD8E6",        # Blue
    "WC UHB": "FFFF99",       # Yellow
    "Annual Leave": "FF6347", # Red
    "Sick Leave": "FF6347",
    "Conference": "FF6347",
    "Training": "FF6347",
    "Personal Leave": "FF6347"
}

# Apply colors
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=ws.max_column):
    for cell in row:
        fill_color = color_map.get(cell.value)
        if fill_color:
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

wb.save(excel_filename)
print(f"Colored rota saved to '{excel_filename}'")


# -------------------------------
# 7️⃣ Print summaries
# -------------------------------
print("Rota saved to 'on_call_rota_16people.xlsx'\n")

print("Weekend summary:")
for p in people:
    total_weekends = weekend_assigned_southmead.get(p,0)+weekend_assigned_uhbw.get(p,0)
    print(f"{p:<8} {fte[p]*100:>3.0f}% FTE  {total_weekends} weekends (Target: {target_weekends[p]:.1f})")

print("\nWeekday summary:")
for p in people:
    weekday_count = sum(1 for d, assigned in rota.items() if assigned==p)
    print(f"{p:<8} {weekday_count} weekdays (Target: {target_shifts[p]:.1f})")