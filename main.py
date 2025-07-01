import yaml
from pathlib import Path
import glob
import os
import pandas as pd

with open("config.yaml", "r") as file:
    config = yaml.safe_load(file)

input_path = config["input_path"]
output_path = config["output_path"]
archive_path = config["archive_path"]

for path in [Path(output_path), Path(archive_path)]:
    if not path.exists():
        path.mkdir(parents=True)

if not Path(input_path).exists():
    raise NotADirectoryError("Input path does not exist.")

excel_files = glob.glob(os.path.join(input_path, "*.xlsx"))
if len(excel_files) != 1:
    raise ValueError("There should be exactly one Excel file in the input path.")

excel_file_path = excel_files[0]

employee_data = pd.read_excel(excel_file_path, sheet_name="Mitarbeiterliste", skiprows=2, header=None, usecols="A:C, E")
special_dates_data = pd.read_excel(excel_file_path, sheet_name="Sondertermine", skiprows=2, header=None)
planning_data = pd.read_excel(excel_file_path, sheet_name="Dienstplanung", header=None)

employee_dict = {
    row[0]: (row[1], row[2])
    for row in employee_data.itertuples(index=False)
}

special_dates_dict = {
    row[0]: (row[1], row[2], row[3], row[4], row[5])
    for row in special_dates_data.itertuples(index=True)
}

possible_assignments = employee_data[employee_data[4].notna()][4].tolist()
possible_groups = possible_assignments[:6]

year = planning_data[1][0]
calendar_week = planning_data[1][1]

start_date = planning_data[1][3].strftime("%Y-%m-%d")
end_date = planning_data[1][5].strftime("%Y-%m-%d")

planning_data_dict = {
    row[0]: (row[1], row[2], row[3], row[4], row[5])
    for row in planning_data.itertuples(index=True)
}

days = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]
cols_per_day = 6
temp_frame = planning_data.iloc[12:]

employee_times = []

for i in range(0, len(temp_frame), 4):
    row_1 = temp_frame.iloc[i]
    row_2 = temp_frame.iloc[i + 1]
    row_3 = temp_frame.iloc[i + 2]
    row_4 = temp_frame.iloc[i + 3]

    employee_name = row_1[0]
    working_times = []

    for day_idx, day in enumerate(days):
        start = (day_idx * cols_per_day) + 2
        end = (start + cols_per_day) - 1

        times_1 = row_1.iloc[start:end].fillna("-").tolist()
        times_2 = row_2.iloc[start:end].fillna("-").tolist()

        working_times.append({
            "day": day,
            "entry_1": {
                "start": times_1[0],
                "end": times_1[1],
                "break_start": times_1[2],
                "break_end": times_1[3],
                "assignment": times_1[4]
            },
            "entry_2": {
                "start": times_2[0],
                "end": times_2[1],
                "break_start": times_2[2],
                "break_end": times_2[3],
                "assignment": times_2[4]
            }
        })

    additional_times = []
    for day_idx, day in enumerate(days):
        start = (day_idx * cols_per_day) + 2
        end = (start + cols_per_day) - 1

        times_1 = row_3.iloc[start:end].fillna("-").tolist()
        times_2 = row_4.iloc[start:end].fillna("-").tolist()

        additional_times.append({
            "day": day,
            "entry_1": {
                "start": times_1[0],
                "end": times_1[1],
                "break_start": times_1[2],
                "break_end": times_1[3],
                "assignment": times_1[4]
            },
            "entry_2": {
                "start": times_2[0],
                "end": times_2[1],
                "break_start": times_2[2],
                "break_end": times_2[3],
                "assignment": times_2[4]
            }
        })

    working_hours_week = round(row_1.iloc[32], 2)
    week_saldo = round(row_1.iloc[33], 2)

    employee_times.append({
        "name": employee_name,
        "working_times": working_times,
        "additional_times": additional_times,
        "working_hours_week": working_hours_week,
        "week_saldo": week_saldo
    })

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import time, datetime, timedelta

data = employee_times
day = "Montag"
pdf_filename = f"{output_path}/dienstplan_{day.lower()}.pdf"

fig, ax = plt.subplots(figsize=(14, 0.8 * len(data)))
yticklabels = []
color_map = {
    "Blau": "#aec6cf",
    "Grün": "#b2d8b2",
    "Rot": "#f4cccc",
    "Nest Gelb": "#fff2b2",
    "Nest Orange": "#ffd8b1",
    "Übergreifend": "#ffe0b3",
    "Verfügungszeit": "#d9d9d9",
    "Elterngespräch": "#d5a6bd",
    "Krank": "#cccccc",
    "-": "white"
}
legend_patches = {}

def time_to_float(t):
    return t.hour + t.minute / 60 if isinstance(t, time) else None

def format_time(t):
    return t.strftime("%H:%M") if isinstance(t, time) else ""

all_times = []
for person in data:
    for block in [person.get("working_times", []), person.get("additional_times", [])]:
        day_data = next((entry for entry in block if entry["day"] == day), None)
        if not day_data:
            continue
        for key in ["entry_1", "entry_2"]:
            entry = day_data.get(key, {})
            start = entry.get("start")
            end = entry.get("end")
            if isinstance(start, time) and isinstance(end, time):
                all_times.append(time_to_float(start))
                all_times.append(time_to_float(end))

start_hour = int(min(all_times)) - 1 if all_times else 6
end_hour = int(max(all_times)) + 1 if all_times else 21

y_spacing = 2
for i, person in enumerate(data):
    y = len(data) * y_spacing - i * y_spacing - 1
    yticklabels.append(person["name"])

    for block in [person.get("working_times", []), person.get("additional_times", [])]:
        day_data = next((entry for entry in block if entry["day"] == day), None)
        if not day_data:
            continue

        for key in ["entry_1", "entry_2"]:
            entry = day_data.get(key, {})
            start = time_to_float(entry.get("start"))
            end = time_to_float(entry.get("end"))
            start_obj = entry.get("start")
            end_obj = entry.get("end")
            assignment = entry.get("assignment", "-")

            if start is None or end is None or assignment == "-":
                continue

            width = end - start
            color = color_map.get(assignment, "#e6e6e6")
            ax.barh(y, width, left=start, height=0.8, color=color, edgecolor="black")

            start_text = format_time(start_obj)
            end_text = format_time(end_obj)

            min_block_duration = 0.15 # entspricht 9 Minuten
            block_width = end - start

            y_base = y - 0.55
            ax.text(start, y_base, start_text, fontsize=5, ha="center", va="top", color="black")

            if block_width < min_block_duration:
                ax.text(end, y_base - 0.3, end_text, fontsize=5, ha="center", va="top", color="black")
            else:
                ax.text(end, y_base, end_text, fontsize=5, ha="center", va="top", color="black")

            if assignment not in legend_patches:
                legend_patches[assignment] = mpatches.Patch(color=color, label=assignment)

ax.set_xlim(start_hour, end_hour)
ax.set_yticks([len(data) * y_spacing - i * y_spacing - 1 for i in range(len(data))])
ax.set_yticklabels(yticklabels)

xticks = list(range(start_hour, end_hour + 1))
xtick_labels = [(datetime(2023, 1, 1, h, 0)).strftime("%H:%M") for h in xticks]
ax.set_xticks(xticks)
ax.set_xticklabels(xtick_labels)

padding_y = 0.5
ylim_lower = -padding_y
ylim_upper = len(data) * y_spacing + padding_y
ax.set_ylim(ylim_lower, ylim_upper)

ax.grid(axis="x", linestyle="--", linewidth=0.5, alpha=0.3)
ax.set_title(f"Dienstplan für {day} den {start_date}", fontsize=14)
ax.legend(handles=legend_patches.values(), bbox_to_anchor=(1.05, 1), loc="upper left")

plt.tight_layout()

plt.savefig(pdf_filename, format="pdf")
print(f"✅ PDF gespeichert unter: {pdf_filename}")