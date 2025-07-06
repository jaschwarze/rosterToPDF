import pandas as pd

def create_time_entry(times_data):
    return {
        "start": times_data[0],
        "end": times_data[1],
        "break_start": times_data[2],
        "break_end": times_data[3],
        "assignment": times_data[4]
    }

def parse_employee_times(frame, cols_per_day, days_of_week):
    employee_times = []
    rows_per_employee = 6

    for i in range(0, len(frame), rows_per_employee):
        rows = [frame.iloc[i + j] for j in range(rows_per_employee)]
        employee_name = rows[0][0]

        if pd.isna(employee_name):
            continue

        def create_times_category(row_indices):
            return [
                {
                    "day": day,
                    "entry_1": create_time_entry(
                        rows[row_indices[0]].iloc[(day_idx * cols_per_day) + 2:(day_idx * cols_per_day) + 2 + cols_per_day - 1].fillna("-").tolist()
                    ),
                    "entry_2": create_time_entry(
                        rows[row_indices[1]].iloc[(day_idx * cols_per_day) + 2:(day_idx * cols_per_day) + 2 + cols_per_day - 1].fillna("-").tolist()
                    )
                }
                for day_idx, day in enumerate(days_of_week)
            ]

        def create_additional_times_category(row_indices):
            return [
                {
                    "day": day,
                    "entry_1": create_time_entry(
                        rows[row_indices[0]].iloc[(day_idx * cols_per_day) + 2:(day_idx * cols_per_day) + 2 + cols_per_day - 1].fillna("-").tolist()
                    ),
                    "entry_2": create_time_entry(
                        rows[row_indices[1]].iloc[(day_idx * cols_per_day) + 2:(day_idx * cols_per_day) + 2 + cols_per_day - 1].fillna("-").tolist()
                    ),
                    "entry_3": create_time_entry(
                        rows[row_indices[2]].iloc[(day_idx * cols_per_day) + 2:(day_idx * cols_per_day) + 2 + cols_per_day - 1].fillna("-").tolist()
                    ),
                    "entry_4": create_time_entry(
                        rows[row_indices[3]].iloc[(day_idx * cols_per_day) + 2:(day_idx * cols_per_day) + 2 + cols_per_day - 1].fillna("-").tolist()
                    )
                }
                for day_idx, day in enumerate(days_of_week)
            ]

        employee_times.append({
            "name": employee_name,
            "working_times": create_times_category([0, 1]),
            "additional_times": create_additional_times_category([2, 3, 4, 5]),
            "working_hours_week": round(rows[0].iloc[32], 2),
            "week_saldo": round(rows[0].iloc[33], 2)
        })

    return employee_times