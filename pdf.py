import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import time, datetime, timedelta
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd


def create_employee_view(employee_times, output_path, assignment_map, year, calendar_week, start_date, days_of_week,
                         special_events=None):

    output_filename = f"{output_path}/Mitarbeiterplan-{year}-KW{calendar_week}.pdf"

    with PdfPages(output_filename) as pdf:
        for day_idx, day in enumerate(days_of_week):
            current_date = (datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)).strftime("%d.%m.%Y")
            current_datetime = datetime.strptime(current_date, "%d.%m.%Y")

            day_special_events = _get_special_events_for_day(special_events, current_datetime)
            _create_employee_view_for_day(pdf, day, employee_times, assignment_map, calendar_week, current_date,
                                          day_special_events)


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


def _get_affected_employees(employee_times, day, assignment, special_start_time, special_end_time):
    affected_employees = []

    special_start_float = _time_to_float(special_start_time)
    special_end_float = _time_to_float(special_end_time)

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

                if (not isinstance(entry_start, time) or not isinstance(entry_end, time) or
                        entry_assignment == "-" or entry_start == "-" or entry_end == "-" or entry_start > entry_end):
                    continue

                entry_start_float = _time_to_float(entry_start)
                entry_end_float = _time_to_float(entry_end)

                time_overlap = (entry_start_float < special_end_float and
                                entry_end_float > special_start_float)

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

    for block_type, block_data in [("working", person.get("working_times", [])),
                                   ("additional", person.get("additional_times", []))]:
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

        ax.plot([x, x], [y_base - 0.1, label_y + 0.05],
                color="black", alpha=0.5, linestyle="--", linewidth=0.5, zorder=1)

        ax.text(x, label_y, text, fontsize=5, ha="center", va="top",
                color="black", alpha=1.0, weight="normal", zorder=2)


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

            if (isinstance(start, time) and isinstance(end, time) and
                    assignment != "-" and start != "-" and end != "-" and start <= end):
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
        time_info = f"{start_time.strftime('%H:%M')}-{end_time.strftime('%H:%M')}"

        legend_handles.append(mpatches.Patch(color="none", label=""))
        legend_labels.append(f"  {event_name} ({time_info})")

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
            all_times.append(_time_to_float(start_time))
            all_times.append(_time_to_float(end_time))

    start_hour = int(min(all_times)) - 1 if all_times else 6
    end_hour = int(max(all_times)) + 1 if all_times else 21

    block_height = 0.9

    y_spacing = _calculate_dynamic_spacing(filtered_data, day)
    for i, person in enumerate(filtered_data):
        y = len(filtered_data) * y_spacing - i * y_spacing - 1
        yticklabels.append(person["name"])

        all_labels = _collect_all_time_labels(person, day)
        for block_type, block_data in [("working", person.get("working_times", [])),
                                       ("additional", person.get("additional_times", []))]:
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