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

def _create_table(ax, table_data, title, fontsize=10, scale=(1.2, 1.2), cell_height=0.05, header_color="#40466e", header_fontcolor="white", color_map=None):
    ax.axis("off")
    table = ax.table(cellText=table_data, cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(fontsize)
    table.scale(*scale)

    for i in range(len(table_data)):
        for j in range(len(table_data[0])):
            cell = table[i, j]
            cell.set_height(cell_height)

            if i == 0 or (j == 0 and i > 0):
                cell.set_text_props(weight="bold", color=header_fontcolor)
                cell.set_facecolor(color_map[i][j] if color_map and i == 0 else header_color)
            else:
                cell.set_facecolor(color_map[i][j] if color_map and i > 0 else "#f7f7f7" if (i + j) % 2 else "#ffffff")

            cell.set_text_props(ha="center", va="center")

    ax.set_title(title, fontsize=14, pad=20)
    plt.tight_layout()

def _create_bar_chart(ax, data, days_of_week, labels, title, width=0.15, colors=None):
    x = np.arange(len(days_of_week))

    for i, label in enumerate(labels):
        values = [data[day][label] for day in days_of_week]
        ax.bar(x + i * width, values, width, label=label, color=colors[i] if colors else None)

    ax.set_xlabel("Tage")
    ax.set_ylabel("Anzahl Mitarbeiter" if "verteilung" in title.lower() else "Arbeitsstunden")
    ax.set_title(title)
    ax.set_xticks(x + width * (len(labels) - 1) / 2)
    ax.set_xticklabels(days_of_week)
    ax.legend()
    plt.tight_layout()

def _calculate_duration(start, end, break_start=None, break_end=None):
    start_dt = datetime.combine(datetime.today(), start)
    end_dt = datetime.combine(datetime.today(), end)
    duration = end_dt - start_dt

    if isinstance(break_start, time) and isinstance(break_end, time):
        break_start_dt = datetime.combine(datetime.today(), break_start)
        break_end_dt = datetime.combine(datetime.today(), break_end)
        duration -= (break_end_dt - break_start_dt)

    return duration.total_seconds() / 3600

def _time_to_float(t):
    return t.hour + t.minute / 60 if isinstance(t, time) else None

def _format_time(t):
    return t.strftime("%H:%M") if isinstance(t, time) else ""

def _duration_to_string(duration):
    hours = duration.seconds // 3600
    minutes = (duration.seconds % 3600) // 60
    return f"{hours}:{minutes:02d}" if minutes > 0 else f"{hours}"

def _get_day_data(person, day, block_key="working_times"):
    return next((entry for entry in person.get(block_key, []) if entry["day"] == day), None)

def create_leader_view(employee_times, output_path, assignment_map, year, calendar_week, days_of_week, possible_groups, employee_dict):
    output_filename = f"{output_path}/Leitungsplan-{year}-KW{calendar_week}.pdf"

    with PdfPages(output_filename) as pdf:
        group_counts = _calculate_group_counts(employee_times, days_of_week, possible_groups)
        fig, ax = plt.subplots(figsize=(10, 3))
        table_data = [[""] + possible_groups] + [[day] + [group_counts[day][group] for group in possible_groups] for day in days_of_week]
        color_map = [["#40466e"] * len(table_data[0]) for _ in table_data]

        for i in range(len(table_data)):
            for j in range(1, len(table_data[0])):
                color_map[i][j] = assignment_map.get(possible_groups[j-1], {"color": "#e6e6e6"})["color"] if i == 0 else "#f7f7f7" if (i + j) % 2 else "#ffffff"

        _create_table(ax, table_data, f"Mitarbeiter pro Gruppe - KW {calendar_week} ({year})", 8, (1.2, 0.8), color_map=color_map)
        pdf.savefig()
        plt.close()

        shift_counts, shift_employees = _calculate_shift_counts(employee_times, days_of_week)
        fig, ax = plt.subplots(figsize=(10, 3))
        shifts = list(shift_counts[days_of_week[0]].keys())
        table_data = [[""] + shifts] + [[day] + [shift_counts[day][shift] for shift in shifts] for day in days_of_week]
        _create_table(ax, table_data, f"Mitarbeiter pro Schicht - KW {calendar_week} ({year})", 8, (1.2, 0.8))
        pdf.savefig()
        plt.close()

        fig, ax = plt.subplots(figsize=(12, 10))
        table_data = [["Tag"] + shifts]
        max_names = max(len(shift_employees[day][shift]) for day in days_of_week for shift in shifts)

        for day in days_of_week:
            row = [day]
            for shift in shifts:
                employees = shift_employees[day][shift]
                row.append("\n".join([", ".join(employees[i:i + 2]) for i in range(0, len(employees), 2)]) or "-")

            table_data.append(row)

        _create_table(ax, table_data, f"Mitarbeiter pro Schicht (Namen) - KW {calendar_week} ({year})", 8, (1.2, 1 + max_names * 0.2), cell_height=0.03 + max_names * 0.1)
        pdf.savefig()
        plt.close()

        saldo_data = _calculate_saldo_data(employee_times)
        fig, ax = plt.subplots(figsize=(7, 6))
        table_data = [["Mitarbeiter", "Wöchentliches Saldo (Std.)", "Status"]] + [[entry["name"], f"{entry['saldo']:.2f}", entry["status"]] for entry in saldo_data]
        _create_table(ax, table_data, f"Überstunden- und Saldoübersicht - KW {calendar_week} ({year})")
        pdf.savefig()
        plt.close()

        absence_data = _calculate_absence_data(employee_times, days_of_week)
        fig, ax = plt.subplots(figsize=(10, 6))
        table_data = [["Tag", "Krank", "Urlaub"]]

        for day in days_of_week:
            krank = "\n".join([", ".join(absence_data[day]["Krank"][i:i + 2]) for i in range(0, len(absence_data[day]["Krank"]), 2)]) or "-"
            urlaub = "\n".join([", ".join(absence_data[day]["Urlaub"][i:i + 2]) for i in range(0, len(absence_data[day]["Urlaub"]), 2)]) or "-"
            table_data.append([day, krank, urlaub])

        _create_table(ax, table_data, f"Abwesenheitsübersicht - KW {calendar_week} ({year})")
        pdf.savefig()
        plt.close()

        fig, ax = plt.subplots(figsize=(10, 6))
        data = np.array([[shift_counts[day][shift] for shift in shifts] for day in days_of_week])
        sns.heatmap(data, annot=True, fmt="d", cmap="YlGnBu", ax=ax, xticklabels=shifts, yticklabels=days_of_week)
        ax.set_title(f"Schichtbesetzung Heatmap - KW {calendar_week} ({year})", fontsize=14)
        ax.set_xlabel("Schichten")
        ax.set_ylabel("Tage")
        plt.tight_layout()
        pdf.savefig()
        plt.close()

        fig, ax = plt.subplots(figsize=(12, 6))
        colors = [assignment_map.get(group, {"color": "#e6e6e6"})["color"] for group in possible_groups]
        _create_bar_chart(ax, group_counts, days_of_week, possible_groups, f"Mitarbeiterverteilung nach Gruppen - KW {calendar_week} ({year})", colors=colors)
        pdf.savefig()
        plt.close()

        fig, ax = plt.subplots(figsize=(12, 6))
        _create_bar_chart(ax, shift_counts, days_of_week, shifts, f"Mitarbeiterverteilung nach Schichten - KW {calendar_week} ({year})")
        pdf.savefig()
        plt.close()

        qualification_hours = _calculate_qualification_hours(employee_times, days_of_week, employee_dict)
        fig, ax = plt.subplots(figsize=(12, 6))
        x = np.arange(len(days_of_week))
        width = 0.35
        ax.bar(x - width / 2, [qualification_hours[day]["Fachkraft"] for day in days_of_week], width, label="Fachkraft", color="#1f77b4")
        ax.bar(x + width / 2, [qualification_hours[day]["Integrationskraft"] for day in days_of_week], width, label="Integrationskraft", color="#ff7f0e")
        ax.set_xlabel("Tage")
        ax.set_ylabel("Arbeitsstunden")
        ax.set_title(f"Arbeitszeitverteilung nach Qualifikation - KW {calendar_week} ({year})")
        ax.set_xticks(x)
        ax.set_xticklabels(days_of_week)
        ax.legend()
        plt.tight_layout()
        pdf.savefig()
        plt.close()

        group_hours = _calculate_group_hours(employee_times, days_of_week, possible_groups)
        fig, ax = plt.subplots(figsize=(12, 6))
        _create_bar_chart(ax, group_hours, days_of_week, possible_groups, f"Arbeitsstunden pro Gruppe - KW {calendar_week} ({year})", colors=colors)
        pdf.savefig()
        plt.close()

    print(f"Leitungsplan erstellt unter: {output_filename}")


def _calculate_group_hours(employee_times, days_of_week, possible_groups):
    group_hours = {day: {group: 0 for group in possible_groups} for day in days_of_week}

    for person in employee_times:
        for day in days_of_week:
            day_data = _get_day_data(person, day)

            if not day_data:
                continue

            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                assignment = entry.get("assignment", "-")
                start = entry.get("start")
                end = entry.get("end")
                break_start = entry.get("break_start")
                break_end = entry.get("break_end")

                if not (isinstance(start, time) and isinstance(end, time) and assignment in possible_groups and assignment not in ["Krank", "Urlaub"]):
                    continue

                group_hours[day][assignment] += _calculate_duration(start, end, break_start, break_end)

    return group_hours

def _calculate_shift_counts(employee_times, days_of_week):
    shift_counts = {day: {shift: 0 for shift in SHIFTS} for day in days_of_week}
    shift_employees = {day: {shift: [] for shift in SHIFTS} for day in days_of_week}

    def time_overlap(entry_start, entry_end, shift_start, shift_end):
        return entry_start < shift_end and entry_end > shift_start

    for person in employee_times:
        for day in days_of_week:
            day_data = _get_day_data(person, day)

            if not day_data:
                continue

            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                start = entry.get("start")
                end = entry.get("end")
                assignment = entry.get("assignment", "-")

                if not (isinstance(start, time) and isinstance(end, time) and assignment not in ["Krank", "Urlaub"]):
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
            day_data = _get_day_data(person, day)

            if not day_data:
                continue

            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                assignment = entry.get("assignment", "-")

                if assignment in possible_groups and assignment not in ["Krank", "Urlaub"] and isinstance(entry.get("start"), time) and isinstance(entry.get("end"), time):
                    group_counts[day][assignment] += 1

    return group_counts

def _calculate_saldo_data(employee_times):
    saldo_data = [{"name": person["name"], "saldo": round(person.get("week_saldo", 0), 2),
                   "status": "Positiv" if person.get("week_saldo", 0) > 0 else "Negativ" if person.get("week_saldo", 0) < 0 else "Neutral"}
                  for person in employee_times]
    return sorted(saldo_data, key=lambda x: x["saldo"], reverse=True)

def _calculate_absence_data(employee_times, days_of_week):
    absence_data = {day: {"Krank": [], "Urlaub": []} for day in days_of_week}
    for person in employee_times:
        for day in days_of_week:
            day_data = _get_day_data(person, day)

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
            day_data = _get_day_data(person, day)

            if not day_data:
                continue

            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                start = entry.get("start")
                end = entry.get("end")
                break_start = entry.get("break_start")
                break_end = entry.get("break_end")
                assignment = entry.get("assignment", "-")

                if not (isinstance(start, time) and isinstance(end, time) and assignment not in ["Krank", "Urlaub"]):
                    continue

                qualification_hours[day][position] += _calculate_duration(start, end, break_start, break_end)

    return qualification_hours

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

    if _check_for_special_events(special_events, assignment, start_date, days_of_week):
        max_counter = max(len(_get_special_events_for_day(special_events, datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)))
                         for day_idx, day in enumerate(days_of_week))
        special_event_height = 0.5 + max_counter * 0.08

    fig, ax = plt.subplots(figsize=(16, max(8, max_employees_per_day * optimal_block_height + 4 + special_event_height)))
    _draw_group_table(ax, group_data, days_of_week, start_date, assignment_map, assignment, special_events, optimal_block_height, special_event_height, employee_dict)
    ax.set_title(f"{'Übergreifend' if assignment == 'Übergreifend' else f'Gruppe: {assignment}'} - KW {calendar_week} ({year})", fontsize=18, fontweight="bold", pad=10)
    ax.set_xlim(0, len(days_of_week))
    ax.set_ylim(0, max_employees_per_day * optimal_block_height + 2 + special_event_height)
    ax.axis("off")
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close()

def _calculate_optimal_block_height(group_data, days_of_week):
    max_text_lines = 0

    for day in days_of_week:
        for employee in group_data[day]:
            target_entries = [e for e in employee["entries"] if e.get("is_target_group", True)]
            main_lines = sum(2 if (e.get("break_start") and e.get("break_end") and isinstance(e["break_start"], time) and isinstance(e["break_end"], time)) else 1 for e in target_entries)
            additional_entries = [e for e in employee["entries"] if not e.get("is_target_group", True)]
            additional_lines = sum(2 if (e.get("break_start") and e.get("break_end") and isinstance(e["break_start"], time) and isinstance(e["break_end"], time)) else 1 for e in additional_entries)
            max_text_lines = max(max_text_lines, main_lines + additional_lines)

    return max(0.6, 0.6 + max_text_lines * 0.08 + 0.05)

def _draw_group_table(ax, group_data, days_of_week, start_date, assignment_map, assignment, special_events, block_height, special_event_height, employee_dict):
    color = assignment_map.get(assignment, {"color": "#e6e6e6"})["color"]
    column_width = 1.0
    max_employees = max(len(group_data[day]) for day in days_of_week)

    for day_idx, day in enumerate(days_of_week):
        x_pos = day_idx
        current_datetime = datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)
        day_special_events = _get_special_events_for_day(special_events, current_datetime)
        fachkraft_duration = timedelta()
        integrationskraft_duration = timedelta()
        special_events_for_assignment = sorted([(event_name, time(0, 0) if pd.isna(start_time) else start_time, time(0, 0) if pd.isna(end_time) else end_time)
                                               for event_id, (event_name, event_date, start_time, end_time, event_assignment) in day_special_events.items()
                                               if event_assignment == assignment or event_assignment == "Übergreifend"], key=lambda x: x[1])

        if special_events_for_assignment and special_event_height > 0:
            gap = 0.1
            special_event_y_pos = max_employees * block_height + 1 + 0.6 + gap
            ax.add_patch(plt.Rectangle((x_pos, special_event_y_pos), column_width, special_event_height, facecolor="#FFCFC2", edgecolor="black", linewidth=1))
            special_event_texts = ["Sondertermine:\n"] + [f"{name}: {start.strftime('%H:%M')} - {end.strftime('%H:%M')} Uhr" if start != time(0, 0) and end != time(0, 0) else
                                                         f"{name}: {start.strftime('%H:%M')} Uhr" if start != time(0, 0) else name for name, start, end in special_events_for_assignment]
            ax.text(x_pos + column_width / 2, special_event_y_pos + special_event_height / 2, "\n".join(special_event_texts), ha="center", va="center", fontsize=9, fontweight="bold", color="black")
        header_y_pos = max_employees * block_height + 1
        current_date = current_datetime.strftime("%d.%m.")
        ax.add_patch(plt.Rectangle((x_pos, header_y_pos), column_width, 0.6, facecolor=color, edgecolor="black", linewidth=1))
        ax.text(x_pos + column_width / 2, header_y_pos + 0.4, day, ha="center", va="center", fontsize=12, fontweight="bold")
        ax.text(x_pos + column_width / 2, header_y_pos + 0.1, current_date, ha="center", va="center", fontsize=10)
        employees = group_data[day]

        for emp_idx, employee in enumerate(employees):
            y_pos = (max_employees - emp_idx - 0.1) * block_height
            ax.add_patch(plt.Rectangle((x_pos, y_pos), column_width, block_height, facecolor="white", edgecolor="black", linewidth=1))
            name_y_pos = y_pos + block_height - 0.15
            ax.text(x_pos + column_width / 2, name_y_pos, employee["name"], ha="center", va="center", fontsize=10, fontweight="bold")
            target_entries = [entry for entry in employee["entries"] if entry.get("is_target_group", True)]
            additional_entries = [entry for entry in employee["entries"] if not entry.get("is_target_group", True)]
            main_time_texts = [f"{entry['start'].strftime('%H:%M')} - {entry['end'].strftime('%H:%M')}\nPause: {entry['break_start'].strftime('%H:%M')}-{entry['break_end'].strftime('%H:%M')}"
                              if entry.get("break_start") and entry.get("break_end") and isinstance(entry["break_start"], time) and isinstance(entry["break_end"], time)
                              else f"{entry['start'].strftime('%H:%M')} - {entry['end'].strftime('%H:%M')}\nPause: ohne" for entry in target_entries]
            additional_time_texts = [f"[{entry['assignment']}] {entry['start'].strftime('%H:%M')} - {entry['end'].strftime('%H:%M')}" +
                                    (f"\nPause: {entry['break_start'].strftime('%H:%M')}-{entry['break_end'].strftime('%H:%M')}" if entry.get("break_start") and entry.get("break_end") and
                                    isinstance(entry["break_start"], time) and isinstance(entry["break_end"], time) else "") for entry in additional_entries]
            total_duration = sum((_calculate_duration(entry["start"], entry["end"], entry.get("break_start"), entry.get("break_end")) for entry in target_entries), 0) * 3600
            main_time_texts.append(f"Arbeitszeit: {_duration_to_string(timedelta(seconds=total_duration))} Std.")
            employee_position = employee_dict.get(employee["name"], {})[1]

            if employee_position == "Fachkraft":
                fachkraft_duration += timedelta(seconds=total_duration)
            elif employee_position == "Integrationskraft":
                integrationskraft_duration += timedelta(seconds=total_duration)
            text_start_y = name_y_pos - 0.15

            if main_time_texts:
                main_text = "\n".join(main_time_texts)
                main_lines = len(main_text.split("\n"))
                ax.text(x_pos + column_width / 2, text_start_y - (main_lines * 0.04), main_text, ha="center", va="center", fontsize=9, color="black")

            if additional_time_texts:
                additional_text = "\n".join(additional_time_texts)
                additional_lines = len(additional_text.split("\n"))
                main_lines = len("\n".join(main_time_texts).split("\n")) if main_time_texts else 0
                ax.text(x_pos + column_width / 2, text_start_y - (main_lines * 0.06) - 0.2 - (additional_lines * 0.04), additional_text, ha="center", va="center", fontsize=8, color="gray", style="italic")

        ax.add_patch(plt.Rectangle((x_pos, 0.25), column_width, 0.4, facecolor="lightgrey", edgecolor="black", linewidth=1.5))
        ax.text(x_pos + column_width / 2, 0.25 + 0.2, f"FK: {_duration_to_string(fachkraft_duration)} Std.\nIK: {_duration_to_string(integrationskraft_duration)} Std.", ha="center", va="center", fontsize=10)

def _collect_group_data(employee_times, target_assignment, days_of_week):
    group_data = {day: [] for day in days_of_week}

    def times_overlap(start1, end1, start2, end2):
        return start1 < end2 and start2 < end1

    for person in employee_times:
        for day in days_of_week:
            target_entries = []
            additional_entries = []

            for block_type, block_data in [("working", person.get("working_times", [])), ("additional", person.get("additional_times", []))]:
                day_data = _get_day_data(person, day, block_type + "_times")

                if not day_data:
                    continue

                for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                    entry = day_data.get(key, {})
                    assignment = entry.get("assignment", "-")
                    start = entry.get("start")
                    end = entry.get("end")

                    if not (isinstance(start, time) and isinstance(end, time) and start <= end and assignment != "-"):
                        continue

                    if assignment == target_assignment and target_assignment not in ["Krank", "Urlaub"]:
                        target_entries.append({"start": start, "end": end, "break_start": entry.get("break_start"), "break_end": entry.get("break_end"), "block_type": block_type, "assignment": assignment, "is_target_group": True})
                    elif assignment != target_assignment and assignment not in ["Krank", "Urlaub", "-"]:
                        for target_entry in target_entries:
                            if times_overlap(start, end, target_entry["start"], target_entry["end"]):
                                additional_entries.append({"start": start, "end": end, "break_start": entry.get("break_start"), "break_end": entry.get("break_end"), "block_type": block_type, "assignment": assignment, "is_target_group": False})
                                break

            if target_entries:
                group_data[day].append({"name": person["name"], "entries": target_entries + additional_entries})
        group_data[day].sort(key=lambda employee: min(entry["start"] for entry in employee["entries"] if entry["is_target_group"]))

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
            _create_employee_view_for_day(pdf, day, employee_times, assignment_map, calendar_week, current_date, _get_special_events_for_day(special_events, current_datetime))

    print(f"Mitarbeiterplan erstellt unter: {output_filename}")

def _get_affected_employees(employee_times, day, assignment, special_start_time, special_end_time):
    affected_employees = []
    special_start_float = _time_to_float(time(0, 0)) if pd.isna(special_start_time) else _time_to_float(special_start_time)
    special_end_float = _time_to_float(time(23, 0)) if pd.isna(special_end_time) else _time_to_float(special_end_time)

    for person in employee_times:
        person_affected = False
        for block in ["working_times", "additional_times"]:
            day_data = _get_day_data(person, day, block)

            if not day_data:
                continue

            for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                entry = day_data.get(key, {})
                entry_assignment = entry.get("assignment", "-")
                entry_start = entry.get("start")
                entry_end = entry.get("end")

                if not (isinstance(entry_start, time) and isinstance(entry_end, time) and entry_assignment != "-" and entry_start <= entry_end):
                    continue

                entry_start_float = _time_to_float(entry_start)
                entry_end_float = _time_to_float(entry_end)

                if entry_start_float < special_end_float and entry_end_float > special_start_float and (assignment == "Übergreifend" or entry_assignment == assignment):
                    person_affected = True
                    break

            if person_affected:
                break

        if person_affected and person["name"] not in affected_employees:
            affected_employees.append(person["name"])

    return affected_employees

def _collect_all_time_labels(person, day):
    labels = []
    seen_times = set()
    for block_type, block_data in [("working", person.get("working_times", [])), ("additional", person.get("additional_times", []))]:
        day_data = _get_day_data(person, day, block_type + "_times")

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

            if start is None or end is None or assignment == "-" or start > end:
                continue

            start_key = (start, _format_time(start_obj))
            end_key = (end, _format_time(end_obj))

            if start_key not in seen_times:
                labels.append({"x": start, "text": _format_time(start_obj), "type": "work_start", "block_type": block_type})
                seen_times.add(start_key)

            if end_key not in seen_times:
                labels.append({"x": end, "text": _format_time(end_obj), "type": "work_end", "block_type": block_type})
                seen_times.add(end_key)

            if break_start is not None and break_end is not None and break_start < break_end:
                break_start_key = (break_start, _format_time(entry.get("break_start")))
                break_end_key = (break_end, _format_time(entry.get("break_end")))

                if break_start_key not in seen_times:
                    labels.append({"x": break_start, "text": _format_time(entry.get("break_start")), "type": "break_start", "block_type": block_type})
                    seen_times.add(break_start_key)

                if break_end_key not in seen_times:
                    labels.append({"x": break_end, "text": _format_time(entry.get("break_end")), "type": "break_end", "block_type": block_type})
                    seen_times.add(break_end_key)

    return labels

def _calculate_label_positions(labels, min_distance=0.3):
    sorted_labels = sorted(labels, key=lambda l: l["x"])

    for label in sorted_labels:
        label["y_level"] = 0
        label["final_x"] = label["x"]

    for i, label in enumerate(sorted_labels):
        conflicts = [sorted_labels[j]["y_level"] for j in range(i) if abs(label["x"] - sorted_labels[j]["final_x"]) < min_distance]
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
    max_levels = [max((label["y_level"] for label in _calculate_label_positions(_collect_all_time_labels(person, day))), default=0) for person in filtered_data]
    return 2.5 + 0.6 * max(max_levels, default=0)

def _has_work_times_for_day(person, day):
    for block in ["working_times", "additional_times"]:
        day_data = _get_day_data(person, day, block)

        if not day_data:
            continue

        for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
            entry = day_data.get(key, {})
            start = entry.get("start")
            end = entry.get("end")
            assignment = entry.get("assignment", "-")

            if isinstance(start, time) and isinstance(end, time) and assignment != "-" and start <= end:
                return True

    return False

def _create_special_events_legend(day_special_events, employee_times, day):
    if not day_special_events:
        return [], []

    legend_handles = [mpatches.Patch(color="none", label=""), mpatches.Patch(color="none", label="Sondertermine:")]
    legend_labels = ["", "Sondertermine:"]

    for event_id, event_data in day_special_events.items():
        event_name, event_date, start_time, end_time, assignment = event_data
        affected_employees = _get_affected_employees(employee_times, day, assignment, start_time, end_time)
        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append(f"  {event_name} ({start_time.strftime('%H:%M')}-{end_time.strftime('%H:%M')})" if not pd.isna(start_time) and not pd.isna(end_time) else f"  {event_name}")
        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append(f"    Gruppe: {assignment}")
        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append(f"    Betroffene Mitarbeiter:")

        if affected_employees:
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
        for block in ["working_times", "additional_times"]:
            day_data = _get_day_data(person, day, block)

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
            day_data = _get_day_data(person, day, block_type + "_times")

            if not day_data:
                continue

            for key in ["entry_1", "entry_2", "entry_3", "entry_4"]:
                entry = day_data.get(key, {})
                start = _time_to_float(entry.get("start"))
                end = _time_to_float(entry.get("end"))
                assignment = entry.get("assignment", "-")
                break_start = _time_to_float(entry.get("break_start"))
                break_end = _time_to_float(entry.get("break_end"))

                if start is None or end is None or assignment == "-" or start > end:
                    continue

                width = end - start
                assignment_entry = assignment_map.get(assignment, {"color": "#e6e6e6", "abbreviation": "?"})
                color = assignment_entry["color"]
                short_label = assignment_entry["abbreviation"]

                if assignment in ["Krank", "Urlaub"]:
                    ax.barh(y, width, left=start, height=block_height, color="#eeeeee", hatch="////", edgecolor="black", alpha=0.6, linewidth=0.2, zorder=3)
                    ax.text(start + width / 2, y, assignment, ha="center", va="center", fontsize=5, zorder=4)
                else:
                    if block_type == "working":
                        ax.barh(y, width, left=start, height=block_height, color=color, edgecolor="black")
                        legend_key = f"{assignment}"

                        if legend_key not in legend_patches:
                            legend_patches[legend_key] = mpatches.Patch(color=color, label=legend_key)

                    elif block_type == "additional":
                        ax.barh(y, width, left=start, height=block_height, color=color, edgecolor="black", alpha=1, linewidth=0.8, linestyle="--")

                        if width > 0.2:
                            ax.text(start + width / 2, y, short_label, ha="center", va="center", fontsize=5, color="black", alpha=0.8)
                        legend_key = f"{assignment}_additional"

                        if legend_key not in legend_patches:
                            legend_patches[legend_key] = mpatches.Patch(color=color, alpha=0.4, label=f"{assignment} ({short_label})", linestyle="--", linewidth=0.8)

                    if break_start is not None and break_end is not None and break_start < break_end:
                        break_width = break_end - break_start
                        ax.barh(y, break_width, left=break_start, height=block_height, color="#eeeeee", hatch="////", edgecolor="black", alpha=0.6, linewidth=0.2, zorder=3)
                        ax.text(break_start + break_width / 2, y, "Pause", ha="center", va="center", fontsize=4, zorder=4)

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
    ax.set_ylim(-padding_y, len(filtered_data) * y_spacing + padding_y)
    ax.grid(axis="x", linestyle="--", linewidth=0.5, alpha=0.3)
    ax.set_title(f"Dienstplan für {day}, den {date} in der KW {calendar_week}", fontsize=14)
    normal_items = [patch for key, patch in legend_patches.items() if not key.endswith("_additional")]
    additional_items = [patch for key, patch in legend_patches.items() if key.endswith("_additional")]
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