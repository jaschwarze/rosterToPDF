import yaml
from pathlib import Path
import glob
import os
import pandas as pd
from datetime import datetime

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

for i in range(0, len(temp_frame), 3):
    row_1 = temp_frame.iloc[i]
    row_2 = temp_frame.iloc[i + 1]
    row_3 = temp_frame.iloc[i + 2]

    employee_name = row_1[0]

    working_times = []

    for day_idx, day in enumerate(days):
        start = day_idx * cols_per_day
        end = start + cols_per_day

        times_1 = row_1.iloc[start:end].tolist()
        times_2 = row_2.iloc[start:end].tolist()

        print(times_1)
        print(times_2)


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