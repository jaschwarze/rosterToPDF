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

employee_data = pd.read_excel(excel_file_path, sheet_name="Mitarbeiterliste", skiprows=2, usecols="A:C, E")
special_dates_data = pd.read_excel(excel_file_path, sheet_name="Sondertermine", skiprows=2)
planning_data = pd.read_excel(excel_file_path, sheet_name="Dienstplanung")


# TODO: Parse Mitarbeiter
# TODO: Parse Gruppen
# TODO: Parse Zurodnungen

print(employee_data)