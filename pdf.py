import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import time, datetime, timedelta
from matplotlib.backends.backend_pdf import PdfPages


def create_employee_view(employee_times, output_path, assignment_map, year, calendar_week, start_date, days_of_week):
    output_filename = f"{output_path}/Mitarbeiterplan-{year}-KW{calendar_week}.pdf"

    with PdfPages(output_filename) as pdf:
        for day_idx, day in enumerate(days_of_week):
            current_date = (datetime.strptime(start_date, "%d.%m.%Y") + timedelta(days=day_idx)).strftime("%d.%m.%Y")
            _create_employee_view_for_day(pdf, day, employee_times, assignment_map, calendar_week, current_date)

def _time_to_float(t):
    return t.hour + t.minute / 60 if isinstance(t, time) else None

def _format_time(t):
    return t.strftime("%H:%M") if isinstance(t, time) else ""

def _create_employee_view_for_day(pdf, day, data, assignment_map, calendar_week, date):
    fig, ax = plt.subplots(figsize=(14, 0.8 * len(data)))
    yticklabels = []
    legend_patches = {}

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
                    all_times.append(_time_to_float(start))
                    all_times.append(_time_to_float(end))

    start_hour = int(min(all_times)) - 1 if all_times else 6
    end_hour = int(max(all_times)) + 1 if all_times else 21

    y_spacing = 2
    for i, person in enumerate(data):
        y = len(data) * y_spacing - i * y_spacing - 1
        yticklabels.append(person["name"])

        for block_type, block_data in [("working", person.get("working_times", [])),
                                       ("additional", person.get("additional_times", []))]:
            day_data = next((entry for entry in block_data if entry["day"] == day), None)
            if not day_data:
                continue

            for key in ["entry_1", "entry_2"]:
                entry = day_data.get(key, {})
                start = _time_to_float(entry.get("start"))
                end = _time_to_float(entry.get("end"))
                start_obj = entry.get("start")
                end_obj = entry.get("end")
                assignment = entry.get("assignment", "-")

                if start is None or end is None or assignment == "-":
                    continue

                width = end - start
                assignment_entry = assignment_map.get(assignment, {"color": "#e6e6e6", "abbr": "?"})
                color = assignment_entry["color"]
                short_label = assignment_entry["abbr"]

                if block_type == "working":
                    ax.barh(y, width, left=start, height=0.8, color=color, edgecolor="black")

                    legend_key = f"{assignment}"
                    if legend_key not in legend_patches:
                        legend_patches[legend_key] = mpatches.Patch(color=color, label=legend_key)

                elif block_type == "additional":
                    ax.barh(y, width, left=start, height=0.8, color=color, edgecolor="black", alpha=0.4, linewidth=0.8,
                            linestyle="--")
                    ax.text(start + width / 2, y, short_label, ha="center", va="center", fontsize=6, color="black",
                            alpha=0.8)

                    legend_key = f"{assignment}_additional"
                    if legend_key not in legend_patches:
                        legend_patches[legend_key] = mpatches.Patch(
                            color=color,
                            alpha=0.4,
                            label=f"{assignment} ({short_label})",
                            linestyle="--",
                            linewidth=0.8
                        )

                start_text = _format_time(start_obj)
                end_text = _format_time(end_obj)
                min_block_duration = 0.15  # entspricht 9 Minuten
                block_width = end - start
                y_base = y - 0.55
                ax.text(start, y_base, start_text, fontsize=5, ha="center", va="top", color="black")

                if block_width < min_block_duration:
                    ax.text(end, y_base - 0.3, end_text, fontsize=5, ha="center", va="top", color="black")
                else:
                    ax.text(end, y_base, end_text, fontsize=5, ha="center", va="top", color="black")

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
    ax.set_title(f"Dienstplan für {day}, den {date}", fontsize=14)

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

    legend = ax.legend(handles=legend_handles, labels=legend_labels, bbox_to_anchor=(1.05, 1), loc="upper left")

    for i, text in enumerate(legend.get_texts()):
        label = text.get_text()
        if label.endswith(":") and not label.startswith("  "):
            text.set_fontweight("bold")
            text.set_fontsize(10)
        elif label.startswith("  "):
            text.set_fontsize(9)
        elif label == "":
            text.set_fontsize(4)

    plt.tight_layout()
    pdf.savefig()
    plt.close()
