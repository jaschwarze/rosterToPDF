import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import time, datetime, timedelta
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import seaborn as sns
import numpy as np

SHIFTS = {
        "Frühdienst": [(time(6, 45), time(7, 0)), (time(7, 0), time(7, 30))],
        "Mittagsdienst": [(time(11, 45), time(13, 30))],
        "Ruhephase 1 \n(12-13 Uhr)": [(time(12, 0), time(13, 0))],
        "Ruhephase 2 \n(13-14 Uhr)": [(time(13, 0), time(14, 0))],
        "Ruhephase 3 \n(14-15 Uhr)": [(time(14, 0), time(15, 0))],
        "Nachmittagsdienst": [(time(15, 0), time(16, 0))]
    }

def create_leader_view(employee_times, output_path, assignment_map, year, calendar_week, days_of_week, possible_groups, employee_dict):
    output_filename = f"{output_path}/Leitungsplan-{year}-KW{calendar_week}.pdf"

    with PdfPages(output_filename) as pdf:
        group_counts = _calculate_group_counts(employee_times, days_of_week, possible_groups)
        _create_group_count_table(pdf, group_counts, days_of_week, possible_groups, year, calendar_week, assignment_map)

        shift_counts, shift_employees = _calculate_shift_counts(employee_times, days_of_week)
        _create_shift_count_table(pdf, shift_counts, days_of_week, year, calendar_week)

        _create_shift_employee_table(pdf, shift_employees, days_of_week, year, calendar_week)

        saldo_data = _calculate_saldo_data(employee_times)
        _create_saldo_table(pdf, saldo_data, year, calendar_week)

        absence_data = _calculate_absence_data(employee_times, days_of_week)
        _create_absence_table(pdf, absence_data, days_of_week, year, calendar_week)

        _create_shift_heatmap(pdf, shift_counts, days_of_week, year, calendar_week)

        _create_group_bar_chart(pdf, group_counts, days_of_week, possible_groups, year, calendar_week, assignment_map)

        _create_shift_bar_chart(pdf, shift_counts, days_of_week, year, calendar_week)

        qualification_hours = _calculate_qualification_hours(employee_times, days_of_week, employee_dict)
        _create_qualification_bar_chart(pdf, qualification_hours, days_of_week, year, calendar_week)

        group_hours = _calculate_group_hours(employee_times, days_of_week, possible_groups)
        _create_group_hours_bar_chart(pdf, group_hours, days_of_week, possible_groups, year, calendar_week, assignment_map)

    print(f"Leitungsplan erstellt unter: {output_filename}")


def _calculate_group_hours(employee_times, days_of_week, possible_groups):
    group_hours = {day: {group: 0 for group in possible_groups} for day in days_of_week}

    for person in employee_times:
        for day in days_of_week:
            day_data = next((entry for entry in person.get("working_times", []) if entry["day"] == day), None)

            if not day_data:
                continue

            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                assignment = entry.get("assignment", "-")
                start = entry.get("start")
                end = entry.get("end")
                break_start = entry.get("break_start")
                break_end = entry.get("break_end")

                if not (isinstance(start, time) and isinstance(end, time) and assignment in possible_groups):
                    continue

                start_dt = datetime.combine(datetime.today(), start)
                end_dt = datetime.combine(datetime.today(), end)
                duration = end_dt - start_dt

                if isinstance(break_start, time) and isinstance(break_end, time):
                    break_start_dt = datetime.combine(datetime.today(), break_start)
                    break_end_dt = datetime.combine(datetime.today(), break_end)
                    duration -= (break_end_dt - break_start_dt)

                hours = duration.total_seconds() / 3600
                group_hours[day][assignment] += hours

    return group_hours


def _create_group_hours_bar_chart(pdf, group_hours, days_of_week, possible_groups, year, calendar_week, assignment_map):
    fig, ax = plt.subplots(figsize=(12, 6))

    x = np.arange(len(days_of_week))
    width = 0.15
    for i, group in enumerate(possible_groups):
        hours = [group_hours[day][group] for day in days_of_week]
        color = assignment_map.get(group, {"color": "#e6e6e6"})["color"]
        ax.bar(x + i * width, hours, width, label=group, color=color)

    ax.set_xlabel("Tage")
    ax.set_ylabel("Arbeitsstunden")
    ax.set_title(f"Arbeitsstunden pro Gruppe - KW {calendar_week} ({year})")
    ax.set_xticks(x + width * (len(possible_groups) - 1) / 2)
    ax.set_xticklabels(days_of_week)
    ax.legend()

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_shift_employee_table(pdf, shift_employees, days_of_week, year, calendar_week):
    fig, ax = plt.subplots(figsize=(12, 10))
    ax.axis("off")

    shifts = list(shift_employees[days_of_week[0]].keys())
    table_data = [["Tag"] + shifts]

    max_names = 1
    for day in days_of_week:
        for shift in shifts:
            num_names = len(shift_employees[day][shift])
            max_names = max(max_names, num_names)

    for day in days_of_week:
        row = [day]
        for shift in shifts:
            employees = shift_employees[day][shift]
            if employees:
                employee_text = ""
                for i in range(0, len(employees), 2):
                    group = employees[i:i + 2]
                    employee_text += ", ".join(group) + "\n"
            else:
                employee_text = "-"
            row.append(employee_text)
        table_data.append(row)


    table = ax.table(cellText=table_data, cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(8)

    table.scale(1.2, 1 + max_names * 0.2)

    for i in range(len(table_data)):
        for j in range(len(table_data[0])):
            cell = table[i, j]
            if i == 0:
                cell.set_text_props(weight="bold", color="white")
                cell.set_height(0.03)
                cell.set_facecolor("#40466e")
            else:
                cell.set_facecolor("#f7f7f7" if i % 2 else "#ffffff")

            cell.set_text_props(ha="center", va="center")
            cell.set_height(cell.get_height() * (1 + max_names * 0.1))

    ax.set_title(f"Mitarbeiter pro Schicht (Namen) - KW {calendar_week} ({year})", fontsize=14, pad=20)

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _calculate_shift_counts(employee_times, days_of_week):
    shift_counts = {day: {shift: 0 for shift in SHIFTS} for day in days_of_week}
    shift_employees = {day: {shift: [] for shift in SHIFTS} for day in days_of_week}

    def time_overlap(entry_start, entry_end, shift_start, shift_end):
        return entry_start < shift_end and entry_end > shift_start

    for person in employee_times:
        for day in days_of_week:
            day_data = next((entry for entry in person.get("working_times", []) if entry["day"] == day), None)
            if not day_data:
                continue
            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                start = entry.get("start")
                end = entry.get("end")
                if not (isinstance(start, time) and isinstance(end, time)):
                    continue
                for shift_name, shift_times in SHIFTS.items():
                    for shift_start, shift_end in shift_times:
                        if time_overlap(start, end, shift_start, shift_end):
                            shift_counts[day][shift_name] += 1
                            if person["name"] not in shift_employees[day][shift_name]:
                                shift_employees[day][shift_name].append(person["name"])
                            break

    return shift_counts, shift_employees

def _calculate_group_counts(employee_times, days_of_week, possible_groups):
    group_counts = {day: {group: 0 for group in possible_groups} for day in days_of_week}

    for person in employee_times:
        for day in days_of_week:
            day_data = next((entry for entry in person.get("working_times", []) if entry["day"] == day), None)
            if not day_data:
                continue
            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                assignment = entry.get("assignment", "-")
                if assignment in possible_groups and isinstance(entry.get("start"), time) and isinstance(
                        entry.get("end"), time):
                    group_counts[day][assignment] += 1

    return group_counts

def _calculate_saldo_data(employee_times):
    saldo_data = []
    for person in employee_times:
        saldo = person.get("week_saldo", 0)
        status = "Positiv" if saldo > 0 else "Negativ" if saldo < 0 else "Neutral"
        saldo_data.append({
            "name": person["name"],
            "saldo": round(saldo, 2),
            "status": status
        })
    return sorted(saldo_data, key=lambda x: x["saldo"], reverse=True)


def _calculate_absence_data(employee_times, days_of_week):
    absence_data = {day: {"Krank": [], "Urlaub": []} for day in days_of_week}

    for person in employee_times:
        for day in days_of_week:
            day_data = next((entry for entry in person.get("working_times", []) if entry["day"] == day), None)
            if not day_data:
                continue
            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                assignment = entry.get("assignment", "-")
                if assignment in ["Krank", "Urlaub"]:
                    absence_data[day][assignment].append(person["name"])

    return absence_data


def _calculate_qualification_hours(employee_times, days_of_week, employee_dict):
    qualification_hours = {day: {"Fachkraft": 0, "Integrationskraft": 0} for day in days_of_week}

    for person in employee_times:
        position = employee_dict.get(person["name"], (None, None))[1]

        if position not in ["Fachkraft", "Integrationskraft"]:
            continue

        for day in days_of_week:
            day_data = next((entry for entry in person.get("working_times", []) if entry["day"] == day), None)

            if not day_data:
                continue

            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                start = entry.get("start")
                end = entry.get("end")
                break_start = entry.get("break_start")
                break_end = entry.get("break_end")

                if not (isinstance(start, time) and isinstance(end, time)):
                    continue

                start_dt = datetime.combine(datetime.today(), start)
                end_dt = datetime.combine(datetime.today(), end)
                duration = end_dt - start_dt

                if isinstance(break_start, time) and isinstance(break_end, time):
                    break_start_dt = datetime.combine(datetime.today(), break_start)
                    break_end_dt = datetime.combine(datetime.today(), break_end)
                    duration -= (break_end_dt - break_start_dt)

                hours = duration.total_seconds() / 3600
                qualification_hours[day][position] += hours

    return qualification_hours


def _create_group_count_table(pdf, group_counts, days_of_week, possible_groups, year, calendar_week, assignment_map):
    fig, ax = plt.subplots(figsize=(10, 3))
    ax.axis("off")

    table_data = [[""] + possible_groups]
    for day in days_of_week:
        row = [day] + [group_counts[day][group] for group in possible_groups]
        table_data.append(row)

    table = ax.table(cellText=table_data, cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1.2, 0.8)

    for i in range(len(table_data)):
        for j in range(len(table_data[0])):
            cell = table[i, j]
            cell.set_height(0.05)
            if i == 0 and j == 0:
                cell.set_text_props(weight="bold", color="white")
                cell.set_facecolor("#40466e")
            elif i == 0:
                group = possible_groups[j - 1]
                color = assignment_map.get(group, {"color": "#e6e6e6"})["color"]
                cell.set_text_props(weight="bold", color="black")
                cell.set_facecolor(color)
            elif j == 0:
                cell.set_text_props(weight="bold", color="white")
                cell.set_facecolor("#40466e")
            else:
                cell.set_facecolor("#f7f7f7" if (i + j) % 2 else "#ffffff")

    ax.set_title(f"Mitarbeiter pro Gruppe - KW {calendar_week} ({year})", fontsize=14, pad=20)

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_shift_count_table(pdf, shift_counts, days_of_week, year, calendar_week):
    fig, ax = plt.subplots(figsize=(10, 3))
    ax.axis("off")

    shifts = list(shift_counts[days_of_week[0]].keys())
    table_data = [[""] + shifts]

    for day in days_of_week:
        row = [day] + [shift_counts[day][shift] for shift in shifts]
        table_data.append(row)

    table = ax.table(cellText=table_data, cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1.2, 0.8)

    for i in range(len(table_data)):
        for j in range(len(table_data[0])):
            cell = table[i, j]
            cell.set_height(0.05)
            if i == 0 or j == 0:
                cell.set_text_props(weight="bold", color="white")
                cell.set_facecolor("#40466e")
            else:
                cell.set_facecolor("#f7f7f7" if (i + j) % 2 else "#ffffff")

    ax.set_title(f"Mitarbeiter pro Schicht - KW {calendar_week} ({year})", fontsize=14, pad=20)

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_saldo_table(pdf, saldo_data, year, calendar_week):
    fig, ax = plt.subplots(figsize=(7, 6))
    ax.axis("off")

    table_data = [["Mitarbeiter", "Wöchentliches Saldo (Std.)", "Status"]]
    for entry in saldo_data:
        table_data.append([entry["name"], f"{entry['saldo']:.2f}", entry["status"]])

    table = ax.table(cellText=table_data, cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.2)

    for i in range(len(table_data)):
        for j in range(len(table_data[0])):
            cell = table[i, j]
            if i == 0:
                cell.set_text_props(weight="bold", color="white")
                cell.set_facecolor("#40466e")
            else:
                cell.set_facecolor("#f7f7f7" if i % 2 else "#ffffff")

    ax.set_title(f"Überstunden- und Saldoübersicht - KW {calendar_week} ({year})", fontsize=14, pad=20)

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_absence_table(pdf, absence_data, days_of_week, year, calendar_week):
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.axis("off")

    table_data = [["Tag", "Krank", "Urlaub"]]
    for day in days_of_week:
        if absence_data[day]["Krank"]:
            krank = ""
            for i in range(0, len(absence_data[day]["Krank"]), 2):
                people = absence_data[day]["Krank"][i:i + 2]
                krank += ", ".join(people) + "\n"
        else:
            krank = "-"

        if absence_data[day]["Urlaub"]:
            urlaub = ""
            for i in range(0, len(absence_data[day]["Urlaub"]), 2):
                people = absence_data[day]["Urlaub"][i:i + 2]
                urlaub += ", ".join(people) + "\n"
        else:
            urlaub = "-"

        table_data.append([day, krank, urlaub])

    table = ax.table(cellText=table_data, cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.2)

    for i in range(len(table_data)):
        for j in range(len(table_data[0])):
            cell = table[i, j]
            if i == 0:
                cell.set_text_props(weight="bold", color="white")
                cell.set_facecolor("#40466e")
            else:
                cell.set_facecolor("#f7f7f7" if i % 2 else "#ffffff")

    ax.set_title(f"Abwesenheitsübersicht - KW {calendar_week} ({year})", fontsize=14, pad=20)

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_shift_heatmap(pdf, shift_counts, days_of_week, year, calendar_week):
    fig, ax = plt.subplots(figsize=(10, 6))

    shifts = list(shift_counts[days_of_week[0]].keys())
    data = np.array([[shift_counts[day][shift] for shift in shifts] for day in days_of_week])

    sns.heatmap(data, annot=True, fmt="d", cmap="YlGnBu", ax=ax,
                xticklabels=shifts, yticklabels=days_of_week)

    ax.set_title(f"Schichtbesetzung Heatmap - KW {calendar_week} ({year})", fontsize=14)
    ax.set_xlabel("Schichten")
    ax.set_ylabel("Tage")

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_group_bar_chart(pdf, group_counts, days_of_week, possible_groups, year, calendar_week, assignment_map):
    fig, ax = plt.subplots(figsize=(12, 6))

    x = np.arange(len(days_of_week))
    width = 0.15
    for i, group in enumerate(possible_groups):
        counts = [group_counts[day][group] for day in days_of_week]
        color = assignment_map.get(group, {"color": "#e6e6e6"})["color"]
        ax.bar(x + i * width, counts, width, label=group, color=color)

    ax.set_xlabel("Tage")
    ax.set_ylabel("Anzahl Mitarbeiter")
    ax.set_title(f"Mitarbeiterverteilung nach Gruppen - KW {calendar_week} ({year})")
    ax.set_xticks(x + width * (len(possible_groups) - 1) / 2)
    ax.set_xticklabels(days_of_week)
    ax.legend()

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_shift_bar_chart(pdf, shift_counts, days_of_week, year, calendar_week):
    fig, ax = plt.subplots(figsize=(12, 6))

    x = np.arange(len(days_of_week))
    width = 0.15
    shifts = list(shift_counts[days_of_week[0]].keys())
    for i, shift in enumerate(shifts):
        counts = [shift_counts[day][shift] for day in days_of_week]
        ax.bar(x + i * width, counts, width, label=shift)

    ax.set_xlabel("Tage")
    ax.set_ylabel("Anzahl Mitarbeiter")
    ax.set_title(f"Mitarbeiterverteilung nach Schichten - KW {calendar_week} ({year})")
    ax.set_xticks(x + width * (len(shifts) - 1) / 2)
    ax.set_xticklabels(days_of_week)
    ax.legend()

    plt.tight_layout()
    pdf.savefig()
    plt.close()


def _create_qualification_bar_chart(pdf, qualification_hours, days_of_week, year, calendar_week):
    fig, ax = plt.subplots(figsize=(12, 6))

    x = np.arange(len(days_of_week))
    width = 0.35

    fachkraft_hours = [qualification_hours[day]["Fachkraft"] for day in days_of_week]
    integrationskraft_hours = [qualification_hours[day]["Integrationskraft"] for day in days_of_week]

    ax.bar(x - width / 2, fachkraft_hours, width, label="Fachkraft", color="#1f77b4")
    ax.bar(x + width / 2, integrationskraft_hours, width, label="Integrationskraft", color="#ff7f0e")

    ax.set_xlabel("Tage")
    ax.set_ylabel("Arbeitsstunden")
    ax.set_title(f"Arbeitszeitverteilung nach Qualifikation - KW {calendar_week} ({year})")
    ax.set_xticks(x)
    ax.set_xticklabels(days_of_week)
    ax.legend()

    plt.tight_layout()
    pdf.savefig()
    plt.close()

def create_group_view(employee_times, output_path, assignment_map, year, calendar_week, start_date, days_of_week, possible_groups, employee_dict, special_events=None):
    output_filename = f"{output_path}/Gruppenplan-{year}-KW{calendar_week}.pdf"

    with PdfPages(output_filename) as pdf:
        for group in possible_groups:
            _create_group_view_for_assignment(pdf, group, employee_times, assignment_map, year, calendar_week, start_date, days_of_week, employee_dict, special_events)

    print(f"Gruppenplan erstellt unter: {output_filename}")


def _create_group_view_for_assignment(pdf, assignment, employee_times, assignment_map, year, calendar_week, start_date, days_of_week, employee_dict, special_events=None):
    group_data = _collect_group_data(employee_times, assignment, days_of_week)
    if not any(group_data[day] for day in days_of_week):
        return

    max_employees_per_day = max(len(group_data[day]) for day in days_of_week)
    optimal_block_height = _calculate_optimal_block_height(group_data, days_of_week)

    special_event_height = 0
    has_special_events = _check_for_special_events(special_events, assignment, start_date, days_of_week)
    if has_special_events:
        max_counter = 0
        for day_idx, day in enumerate(days_of_week):
            current_datetime = datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)
            day_special_events = _get_special_events_for_day(special_events, current_datetime)
            if len(day_special_events) > max_counter:
                max_counter = len(day_special_events)

        special_event_height = 0.5 + max_counter * 0.08

    fig, ax = plt.subplots(figsize=(16, max(8, max_employees_per_day * optimal_block_height + 4 + special_event_height)))

    _draw_group_table(ax, group_data, days_of_week, start_date, assignment_map, assignment, special_events, optimal_block_height, special_event_height, employee_dict)

    if assignment == "Übergreifend":
        title = f"{assignment} - KW {calendar_week} ({year})"
    else:
        title = f"Gruppe: {assignment} - KW {calendar_week} ({year})"

    ax.set_title(title, fontsize=18, fontweight="bold", pad=10)

    ax.set_xlim(0, len(days_of_week))
    ax.set_ylim(0, max_employees_per_day * optimal_block_height + 2 + special_event_height)
    ax.axis("off")

    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close()


def _calculate_optimal_block_height(group_data, days_of_week):
    max_text_lines = 0

    for day_idx, day in enumerate(days_of_week):
        employees = group_data[day]

        for employee in employees:
            target_entries = [entry for entry in employee["entries"] if entry.get("is_target_group", True)]
            main_lines = 0
            for entry in target_entries:
                main_lines += 1
                if (entry.get("break_start") and entry.get("break_end") and
                        isinstance(entry["break_start"], time) and isinstance(entry["break_end"], time)):
                    main_lines += 1
                else:
                    main_lines += 1

            additional_entries = [entry for entry in employee["entries"] if not entry.get("is_target_group", True)]
            additional_lines = 0
            for entry in additional_entries:
                additional_lines += 1
                if entry.get("break_start") and entry.get("break_end") and isinstance(entry["break_start"], time) and isinstance(entry["break_end"], time):
                    additional_lines += 1

            total_lines = main_lines + additional_lines
            max_text_lines = max(max_text_lines, total_lines)

    base_height = 0.6
    text_height = max_text_lines * 0.08
    padding = 0.05

    return max(base_height, base_height + text_height + padding)


def _draw_group_table(ax, group_data, days_of_week, start_date, assignment_map, assignment, special_events, block_height, special_event_height, employee_dict):
    assignment_info = assignment_map.get(assignment, {"color": "#e6e6e6"})
    color = assignment_info["color"]

    column_width = 1.0
    max_employees = max(len(group_data[day]) for day in days_of_week)

    for day_idx, day in enumerate(days_of_week):
        x_pos = day_idx

        current_datetime = datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)
        day_special_events = _get_special_events_for_day(special_events, current_datetime)

        fachkraft_duration = timedelta()
        integrationskraft_duration = timedelta()

        special_events_for_assignment = []
        for event_id, event_data in day_special_events.items():
            event_name, event_date, start_time, end_time, event_assignment = event_data
            if event_assignment == assignment or event_assignment == "Übergreifend":
                start_time = time(0, 0) if pd.isna(start_time) else start_time
                end_time = time(0, 0) if pd.isna(end_time) else end_time
                special_events_for_assignment.append((event_name, start_time, end_time))

        special_events_for_assignment.sort(key=lambda x: x[1])

        if special_events_for_assignment and special_event_height > 0:
            gap = 0.1
            special_event_y_pos = max_employees * block_height + 1 + 0.6 + gap
            special_event_rect = plt.Rectangle((x_pos, special_event_y_pos), column_width, special_event_height, facecolor="#FFCFC2", edgecolor="black", linewidth=1)
            ax.add_patch(special_event_rect)


            special_event_texts = ["Sondertermine:\n"]
            for event_name, start_time, end_time in special_events_for_assignment:
                time_str = f"{event_name}"
                if isinstance(start_time, time) and isinstance(end_time, time) and start_time != time(0, 0) and end_time != time(0, 0):
                    time_str = f"{event_name}: {start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')} Uhr"
                elif isinstance(start_time, time) and start_time != time(0, 0):
                    time_str = f"{event_name}: {start_time.strftime('%H:%M')} Uhr"

                special_event_texts.append(time_str)

            special_event_text = "\n".join(special_event_texts)
            ax.text(x_pos + column_width / 2, special_event_y_pos + special_event_height / 2, special_event_text, ha="center", va="center", fontsize=9, fontweight="bold", color="black")

        header_y_pos = max_employees * block_height + 1
        current_date = (datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)).strftime("%d.%m.")
        header_rect = plt.Rectangle((x_pos, header_y_pos), column_width, 0.6, facecolor=color, edgecolor="black", linewidth=1)
        ax.add_patch(header_rect)

        ax.text(x_pos + column_width / 2, header_y_pos + 0.4, day, ha="center", va="center", fontsize=12, fontweight="bold")
        ax.text(x_pos + column_width / 2, header_y_pos + 0.1, current_date, ha="center", va="center", fontsize=10)

        employees = group_data[day]
        for emp_idx, employee in enumerate(employees):
            y_pos = (max_employees - emp_idx - 0.1) * block_height

            emp_rect = plt.Rectangle((x_pos, y_pos), column_width, block_height, facecolor="white", edgecolor="black", linewidth=1)
            ax.add_patch(emp_rect)

            name_y_pos = y_pos + block_height - 0.15
            ax.text(x_pos + column_width / 2, name_y_pos, employee["name"], ha="center", va="center", fontsize=10, fontweight="bold")

            target_entries = [entry for entry in employee["entries"] if entry.get("is_target_group", True)]
            additional_entries = [entry for entry in employee["entries"] if not entry.get("is_target_group", True)]

            main_time_texts = []
            for entry in target_entries:
                start_str = entry["start"].strftime("%H:%M")
                end_str = entry["end"].strftime("%H:%M")
                time_text = f"{start_str} - {end_str}"

                if entry.get("break_start") and entry.get("break_end") and isinstance(entry["break_start"], time) and isinstance(entry["break_end"], time):
                    break_start_str = entry["break_start"].strftime("%H:%M")
                    break_end_str = entry["break_end"].strftime("%H:%M")
                    time_text += f"\nPause: {break_start_str}-{break_end_str}"
                else:
                    time_text += "\nPause: ohne"

                main_time_texts.append(time_text)

            additional_time_texts = []
            for entry in additional_entries:
                start_str = entry["start"].strftime("%H:%M")
                end_str = entry["end"].strftime("%H:%M")
                assignment_name = entry.get("assignment", "Unbekannt")
                time_text = f"[{assignment_name}] {start_str} - {end_str}"

                if entry.get("break_start") and entry.get("break_end") and isinstance(entry["break_start"], time) and isinstance(entry["break_end"], time):
                    break_start_str = entry["break_start"].strftime("%H:%M")
                    break_end_str = entry["break_end"].strftime("%H:%M")
                    time_text += f"\nPause: {break_start_str}-{break_end_str}"

                additional_time_texts.append(time_text)

            total_duration = timedelta()
            for entry in target_entries:
                start_datetime = datetime.combine(datetime.today(), entry["start"])
                end_datetime = datetime.combine(datetime.today(), entry["end"])

                if entry.get("break_start") and entry.get("break_end") and isinstance(entry["break_start"], time) and isinstance(entry["break_end"], time):
                    break_start_datetime = datetime.combine(datetime.today(), entry["break_start"])
                    break_end_datetime = datetime.combine(datetime.today(), entry["break_end"])
                    break_duration = break_end_datetime - break_start_datetime
                    total_duration += (end_datetime - start_datetime) - break_duration
                else:
                    total_duration += end_datetime - start_datetime

            main_time_texts.append(f"Arbeitszeit: {_duration_to_string(total_duration)} Std.")

            employee_position = employee_dict.get(employee["name"], {})[1]
            if employee_position == "Fachkraft":
                fachkraft_duration += total_duration
            elif employee_position == "Integrationskraft":
                integrationskraft_duration += total_duration

            text_start_y = name_y_pos - 0.15
            if main_time_texts:
                main_text = "\n".join(main_time_texts)
                main_lines = len(main_text.split("\n"))
                main_text_y = text_start_y - (main_lines * 0.04)

                ax.text(x_pos + column_width / 2, main_text_y, main_text, ha="center", va="center", fontsize=9, color="black")

            if additional_time_texts:
                additional_text = "\n".join(additional_time_texts)
                additional_lines = len(additional_text.split("\n"))

                main_lines = len("\n".join(main_time_texts).split("\n")) if main_time_texts else 0
                additional_text_y = text_start_y - (main_lines * 0.06) - 0.2 - (additional_lines * 0.04)

                ax.text(x_pos + column_width / 2, additional_text_y, additional_text, ha="center", va="center", fontsize=8, color="gray", style="italic")

        bottom_rect = plt.Rectangle((x_pos, 0.25), column_width, 0.4, facecolor="lightgrey", edgecolor="black", linewidth=1.5)
        ax.add_patch(bottom_rect)

        info_text = f"FK: {_duration_to_string(fachkraft_duration)} Std.\nIK: {_duration_to_string(integrationskraft_duration)} Std."
        ax.text(x_pos + column_width / 2, 0.25 + 0.2, info_text, ha="center", va="center", fontsize=10)

def _duration_to_string(duration):
    hours = duration.seconds // 3600
    minutes = (duration.seconds % 3600) // 60
    if minutes > 0:
        duration_text = f"{hours}:{minutes}"
    else:
        duration_text = f"{hours}"

    return duration_text

def _collect_group_data(employee_times, target_assignment, days_of_week):
    group_data = {day: [] for day in days_of_week}

    def times_overlap(start1, end1, start2, end2):
        return start1 < end2 and start2 < end1

    def time_within_range(check_start, check_end, range_start, range_end):
        return times_overlap(check_start, check_end, range_start, range_end)

    for person in employee_times:
        for day_idx, day in enumerate(days_of_week):
            target_entries = []

            for block_type, block_data in [("working", person.get("working_times", [])), ("additional", person.get("additional_times", []))]:
                day_data = next((entry for entry in block_data if entry["day"] == day), None)
                if not day_data:
                    continue

                for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                    entry = day_data.get(key, {})
                    assignment = entry.get("assignment", "-")
                    start = entry.get("start")
                    end = entry.get("end")

                    if assignment == target_assignment and target_assignment != "Krank" and target_assignment != "Urlaub" and isinstance(start, time) and isinstance(end, time) and start != "-" and end != "-" and start <= end:
                        target_entries.append({
                            "start": start,
                            "end": end,
                            "break_start": entry.get("break_start"),
                            "break_end": entry.get("break_end"),
                            "block_type": block_type,
                            "assignment": assignment,
                            "is_target_group": True
                        })

            if not target_entries:
                continue

            additional_entries = []

            for block_type, block_data in [("working", person.get("working_times", [])), ("additional", person.get("additional_times", []))]:
                day_data = next((entry for entry in block_data if entry["day"] == day), None)
                if not day_data:
                    continue

                for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                    entry = day_data.get(key, {})
                    assignment = entry.get("assignment", "-")
                    start = entry.get("start")
                    end = entry.get("end")

                    if assignment != target_assignment and assignment not in ["Krank", "Urlaub", "-"] and isinstance(start, time) and isinstance(end, time) and start != "-" and end != "-" and start <= end:
                        for target_entry in target_entries:
                            if time_within_range(start, end, target_entry["start"], target_entry["end"]):
                                additional_entries.append({
                                    "start": start,
                                    "end": end,
                                    "break_start": entry.get("break_start"),
                                    "break_end": entry.get("break_end"),
                                    "block_type": block_type,
                                    "assignment": assignment,
                                    "is_target_group": False
                                })
                                break

            all_entries = target_entries + additional_entries

            if all_entries:
                group_data[day].append({
                    "name": person["name"],
                    "entries": all_entries
                })

    for day in days_of_week:
        group_data[day].sort(
            key=lambda employee: min(entry["start"] for entry in employee["entries"] if entry["is_target_group"]))

    return group_data


def _check_for_special_events(special_events, assignment, start_date, days_of_week):
    if not special_events:
        return False

    for day_idx, day in enumerate(days_of_week):
        current_datetime = datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)
        day_special_events = _get_special_events_for_day(special_events, current_datetime)

        for event_id, event_data in day_special_events.items():
            event_name, event_date, start_time, end_time, event_assignment = event_data
            if event_assignment == assignment or event_assignment == "Übergreifend":
                return True

    return False

def _get_special_events_for_day(special_events, target_date):
    if not special_events:
        return {}

    day_events = {}
    for event_id, event_data in special_events.items():
        event_name, event_date, start_time, end_time, assignment = event_data

        if hasattr(event_date, "to_pydatetime"):
            event_date = event_date.to_pydatetime()
        elif isinstance(event_date, pd.Timestamp):
            event_date = event_date.to_pydatetime()

        if event_date.date() == target_date.date():
            day_events[event_id] = event_data

    return day_events


def create_employee_view(employee_times, output_path, assignment_map, year, calendar_week, start_date, days_of_week, special_events=None):
    output_filename = f"{output_path}/Mitarbeiterplan-{year}-KW{calendar_week}.pdf"

    with PdfPages(output_filename) as pdf:
        for day_idx, day in enumerate(days_of_week):
            current_date = (datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)).strftime("%d.%m.%Y")
            current_datetime = datetime.strptime(current_date, "%d.%m.%Y")

            day_special_events = _get_special_events_for_day(special_events, current_datetime)
            _create_employee_view_for_day(pdf, day, employee_times, assignment_map, calendar_week, current_date, day_special_events)

    print(f"Mitarbeiterplan erstellt unter: {output_filename}")

def _get_affected_employees(employee_times, day, assignment, special_start_time, special_end_time):
    affected_employees = []
    special_start_float = _time_to_float(time(0, 0)) if pd.isna(special_start_time) else _time_to_float(special_start_time)
    special_end_float = _time_to_float(time(23, 0)) if pd.isna(special_end_time) else _time_to_float(special_end_time)

    for person in employee_times:
        if not _has_work_times_for_day(person, day):
            continue

        person_affected = False

        for block in [person.get("working_times", []), person.get("additional_times", [])]:
            day_data = next((entry for entry in block if entry["day"] == day), None)
            if not day_data:
                continue

            for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                entry = day_data.get(key, {})
                entry_assignment = entry.get("assignment", "-")
                entry_start = entry.get("start")
                entry_end = entry.get("end")

                if not isinstance(entry_start, time) or not isinstance(entry_end, time) or entry_assignment == "-" or entry_start == "-" or entry_end == "-" or entry_start > entry_end:
                    continue

                entry_start_float = _time_to_float(entry_start)
                entry_end_float = _time_to_float(entry_end)

                time_overlap = (entry_start_float < special_end_float and entry_end_float > special_start_float)

                if not time_overlap:
                    continue

                if assignment == "Übergreifend":
                    person_affected = True
                    break
                elif entry_assignment == assignment:
                    person_affected = True
                    break

            if person_affected:
                break

        if person_affected and person["name"] not in affected_employees:
            affected_employees.append(person["name"])

    return affected_employees


def _time_to_float(t):
    return t.hour + t.minute / 60 if isinstance(t, time) else None

def _format_time(t):
    return t.strftime("%H:%M") if isinstance(t, time) else ""

def _collect_all_time_labels(person, day):
    labels = []
    seen_times = set()

    for block_type, block_data in [("working", person.get("working_times", [])), ("additional", person.get("additional_times", []))]:
        day_data = next((entry for entry in block_data if entry["day"] == day), None)
        if not day_data:
            continue

        for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
            entry = day_data.get(key, {})
            start = _time_to_float(entry.get("start"))
            end = _time_to_float(entry.get("end"))
            start_obj = entry.get("start")
            end_obj = entry.get("end")
            assignment = entry.get("assignment", "-")
            break_start = _time_to_float(entry.get("break_start"))
            break_end = _time_to_float(entry.get("break_end"))

            if start is None or end is None or assignment == "-" or start > end or start == "-" or end == "-":
                continue

            start_key = (start, _format_time(start_obj))
            end_key = (end, _format_time(end_obj))

            if start_key not in seen_times:
                labels.append({
                    "x": start,
                    "text": _format_time(start_obj),
                    "type": "work_start",
                    "block_type": block_type
                })
                seen_times.add(start_key)

            if end_key not in seen_times:
                labels.append({
                    "x": end,
                    "text": _format_time(end_obj),
                    "type": "work_end",
                    "block_type": block_type
                })
                seen_times.add(end_key)

            if break_start is not None and break_end is not None and break_start < break_end:
                break_start_key = (break_start, _format_time(entry.get("break_start")))
                break_end_key = (break_end, _format_time(entry.get("break_end")))

                if break_start_key not in seen_times:
                    labels.append({
                        "x": break_start,
                        "text": _format_time(entry.get("break_start")),
                        "type": "break_start",
                        "block_type": block_type
                    })
                    seen_times.add(break_start_key)

                if break_end_key not in seen_times:
                    labels.append({
                        "x": break_end,
                        "text": _format_time(entry.get("break_end")),
                        "type": "break_end",
                        "block_type": block_type
                    })
                    seen_times.add(break_end_key)

    return labels


def _calculate_label_positions(labels, min_distance=0.3):
    if not labels:
        return []

    sorted_labels = sorted(labels, key=lambda l: l["x"])

    for label in sorted_labels:
        label["y_level"] = 0
        label["final_x"] = label["x"]

    for i, label in enumerate(sorted_labels):
        conflicts = []

        for j in range(i):
            other = sorted_labels[j]
            distance = abs(label["x"] - other["final_x"])

            if distance < min_distance:
                conflicts.append(other["y_level"])

        level = 0
        while level in conflicts:
            level += 1

        label["y_level"] = level

    return sorted_labels


def _draw_time_labels_with_lines(ax, labels, y_base, base_offset=0.3, level_offset=0.65):
    positioned_labels = _calculate_label_positions(labels)

    for label in positioned_labels:
        x = label["final_x"]
        y_level = label["y_level"]
        text = label["text"]
        label_y = y_base - base_offset - (y_level * level_offset)

        ax.plot([x, x], [y_base - 0.1, label_y + 0.05], color="black", alpha=0.5, linestyle="--", linewidth=0.5, zorder=1)
        ax.text(x, label_y, text, fontsize=5, ha="center", va="top", color="black", alpha=1.0, weight="normal", zorder=2)


def _calculate_dynamic_spacing(filtered_data, day):
    max_levels = []

    for person in filtered_data:
        labels = _collect_all_time_labels(person, day)
        positioned_labels = _calculate_label_positions(labels)

        if positioned_labels:
            max_level = max(label["y_level"] for label in positioned_labels)
            max_levels.append(max_level)

    if not max_levels:
        return 3

    overall_max_level = max(max_levels)
    base_spacing = 2.5
    additional_spacing = 0.6 * overall_max_level

    return base_spacing + additional_spacing

def _has_work_times_for_day(person, day):
    for block in [person.get("working_times", []), person.get("additional_times", [])]:
        day_data = next((entry for entry in block if entry["day"] == day), None)
        if not day_data:
            continue

        for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
            entry = day_data.get(key, {})
            start = entry.get("start")
            end = entry.get("end")
            assignment = entry.get("assignment", "-")

            if isinstance(start, time) and isinstance(end, time) and assignment != "-" and start != "-" and end != "-" and start <= end:
                return True

    return False


def _create_special_events_legend(day_special_events, employee_times, day):
    if not day_special_events:
        return [], []

    legend_handles = []
    legend_labels = []

    legend_handles.append(mpatches.Patch(color="none", label=""))
    legend_labels.append("")
    legend_handles.append(mpatches.Patch(color="none", label=""))
    legend_labels.append("Sondertermine:")

    for event_id, event_data in day_special_events.items():
        event_name, event_date, start_time, end_time, assignment = event_data
        affected_employees = _get_affected_employees(employee_times, day, assignment, start_time, end_time)

        legend_handles.append(mpatches.Patch(color="none", label=""))
        if not pd.isna(start_time) and not pd.isna(end_time):
            time_info = f"{start_time.strftime('%H:%M')}-{end_time.strftime('%H:%M')}"
            legend_labels.append(f"  {event_name} ({time_info})")
        else:
            legend_labels.append(f"  {event_name}")

        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append(f"    Gruppe: {assignment}")

        if affected_employees:
            legend_handles.append(mpatches.Patch(color="none", label=""))
            legend_labels.append(f"    Betroffene Mitarbeiter:")

            for employee in affected_employees:
                legend_handles.append(mpatches.Patch(color="none", label=""))
                legend_labels.append(f"      • {employee}")
        else:
            legend_handles.append(mpatches.Patch(color="none", label=""))
            legend_labels.append("    Keine betroffenen Mitarbeiter")

        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append("")

    return legend_handles, legend_labels


def _create_employee_view_for_day(pdf, day, data, assignment_map, calendar_week, date, day_special_events=None):
    filtered_data = [person for person in data if _has_work_times_for_day(person, day)]

    if not filtered_data:
        return

    fig, ax = plt.subplots(figsize=(16, 0.7 * len(filtered_data)))
    yticklabels = []
    legend_patches = {}

    default_start_hour = 6
    default_end_hour = 21

    all_times = []
    for person in filtered_data:
        for block in [person.get("working_times", []), person.get("additional_times", [])]:
            day_data = next((entry for entry in block if entry["day"] == day), None)
            if not day_data:
                continue
            for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                entry = day_data.get(key, {})
                start = entry.get("start")
                end = entry.get("end")
                break_start = entry.get("break_start")
                break_end = entry.get("break_end")

                if isinstance(start, time) and isinstance(end, time):
                    all_times.append(_time_to_float(start))
                    all_times.append(_time_to_float(end))

                if isinstance(break_start, time) and isinstance(break_end, time):
                    all_times.append(_time_to_float(break_start))
                    all_times.append(_time_to_float(break_end))

    if day_special_events:
        for event_id, event_data in day_special_events.items():
            event_name, event_date, start_time, end_time, assignment = event_data
            start_time = time(default_start_hour, 0) if pd.isna(start_time) else start_time
            end_time = time(default_end_hour, 0) if pd.isna(end_time) else end_time
            all_times.append(_time_to_float(start_time))
            all_times.append(_time_to_float(end_time))

    start_hour = int(min(all_times)) - 1 if all_times else default_start_hour
    end_hour = int(max(all_times)) + 1 if all_times else default_end_hour

    block_height = 0.9

    y_spacing = _calculate_dynamic_spacing(filtered_data, day)
    for i, person in enumerate(filtered_data):
        y = len(filtered_data) * y_spacing - i * y_spacing - 1
        yticklabels.append(person["name"])

        all_labels = _collect_all_time_labels(person, day)
        for block_type, block_data in [("working", person.get("working_times", [])), ("additional", person.get("additional_times", []))]:
            day_data = next((entry for entry in block_data if entry["day"] == day), None)
            if not day_data:
                continue

            for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                entry = day_data.get(key, {})
                start = _time_to_float(entry.get("start"))
                end = _time_to_float(entry.get("end"))
                assignment = entry.get("assignment", "-")
                break_start = _time_to_float(entry.get("break_start"))
                break_end = _time_to_float(entry.get("break_end"))

                if start is None or end is None or assignment == "-" or start > end or start == "-" or end == "-":
                    continue

                width = end - start
                assignment_entry = assignment_map.get(assignment, {"color": "#e6e6e6", "abbreviation": "?"})
                color = assignment_entry["color"]
                short_label = assignment_entry["abbreviation"]

                if assignment == "Krank" or assignment == "Urlaub":
                    ax.barh(
                        y,
                        width,
                        left=start,
                        height=block_height,
                        color="#eeeeee",
                        hatch="////",
                        edgecolor="black",
                        alpha=0.6,
                        linewidth=0.2,
                        zorder=3
                    )
                    ax.text(
                        start + width / 2,
                        y,
                        assignment,
                        ha="center",
                        va="center",
                        fontsize=5,
                        zorder=4
                    )
                else:
                    if block_type == "working":
                        ax.barh(y, width, left=start, height=block_height, color=color, edgecolor="black")
                        legend_key = f"{assignment}"
                        if legend_key not in legend_patches:
                            legend_patches[legend_key] = mpatches.Patch(color=color, label=legend_key)

                    elif block_type == "additional":
                        ax.barh(
                            y,
                            width,
                            left=start,
                            height=block_height,
                            color=color,
                            edgecolor="black",
                            alpha=1,
                            linewidth=0.8,
                            linestyle="--"
                        )

                        if width > 0.2:
                            ax.text(
                                start + width / 2,
                                y,
                                short_label,
                                ha="center",
                                va="center",
                                fontsize=5,
                                color="black",
                                alpha=0.8
                            )

                        legend_key = f"{assignment}_additional"
                        if legend_key not in legend_patches:
                            legend_patches[legend_key] = mpatches.Patch(
                                color=color,
                                alpha=0.4,
                                label=f"{assignment} ({short_label})",
                                linestyle="--",
                                linewidth=0.8
                            )

                    if break_start is not None and break_end is not None and break_start < break_end:
                        break_width = break_end - break_start
                        ax.barh(
                            y,
                            break_width,
                            left=break_start,
                            height=block_height,
                            color="#eeeeee",
                            hatch="////",
                            edgecolor="black",
                            alpha=0.6,
                            linewidth=0.2,
                            zorder=3
                        )
                        ax.text(
                            break_start + break_width / 2,
                            y,
                            "Pause",
                            ha="center",
                            va="center",
                            fontsize=4,
                            zorder=4
                        )

                    y_base = y - 0.65
                    _draw_time_labels_with_lines(ax, all_labels, y_base)

    ax.set_xlim(start_hour, end_hour)
    ax.set_yticks([len(filtered_data) * y_spacing - i * y_spacing - 1 for i in range(len(filtered_data))])
    ax.set_yticklabels(yticklabels)

    xticks = list(range(start_hour, end_hour + 1))
    xtick_labels = [(datetime(2023, 1, 1, h, 0)).strftime("%H:%M") for h in xticks]
    ax.set_xticks(xticks)
    ax.set_xticklabels(xtick_labels)

    padding_y = 0.5 + (y_spacing - 2.5) * 0.3
    ylim_lower = -padding_y
    ylim_upper = len(filtered_data) * y_spacing + padding_y
    ax.set_ylim(ylim_lower, ylim_upper)

    ax.grid(axis="x", linestyle="--", linewidth=0.5, alpha=0.3)
    ax.set_title(f"Dienstplan für {day}, den {date} in der KW {calendar_week}", fontsize=14)

    normal_items = []
    additional_items = []

    for key, patch in legend_patches.items():
        if key.endswith("_additional"):
            additional_items.append(patch)
        else:
            normal_items.append(patch)

    legend_handles = []
    legend_labels = []

    if normal_items:
        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append("Zuweisungen:")

        for item in normal_items:
            legend_handles.append(item)
            legend_labels.append(f"  {item.get_label()}")

    if additional_items:
        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append("")

        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append("Zusätze:")

        for item in additional_items:
            legend_handles.append(item)
            legend_labels.append(f"  {item.get_label()}")

    special_handles, special_labels = _create_special_events_legend(day_special_events, filtered_data, day)
    if special_handles:
        legend_handles.extend(special_handles)
        legend_labels.extend(special_labels)

    legend = ax.legend(handles=legend_handles, labels=legend_labels, bbox_to_anchor=(1.05, 1), loc="upper left")

    for i, text in enumerate(legend.get_texts()):
        label = text.get_text()
        if label.endswith(":") and not label.startswith("  "):
            text.set_fontweight("bold")
            text.set_fontsize(10)
        elif label.startswith("  ") and not label.startswith("    "):
            text.set_fontweight("roman")
            text.set_fontsize(9)
        elif label.startswith("    "):
            text.set_fontsize(8)
        elif label.startswith("      "):
            text.set_fontsize(7)
        elif label == "":
            text.set_fontsize(4)

    plt.tight_layout()
    pdf.savefig()
    plt.close()