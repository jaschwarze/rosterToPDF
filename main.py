import yaml
from pathlib import Path
import glob
import os
import pandas as pd
from pdf import create_employee_view
from parser import parse_employee_times


with open("config.yaml", "r") as file:
    config = yaml.safe_load(file)

input_path = config["input_path"]
output_path = config["output_path"]
archive_path = config["archive_path"]
assignment_map = config.get("assignments", {})
cols_per_day = config.get("cols_per_day", 6)
days_of_week = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]

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
start_date = planning_data[1][3].strftime("%d.%m.%Y")
end_date = planning_data[1][5].strftime("%d.%m.%Y")

planning_data_dict = {
    row[0]: (row[1], row[2], row[3], row[4], row[5])
    for row in planning_data.itertuples(index=True)
}

planning_frame = planning_data.iloc[12:]
employee_times = parse_employee_times(planning_frame, cols_per_day, days_of_week)

create_employee_view(employee_times, output_path, assignment_map, year, calendar_week, start_date, days_of_week)