import re
import datetime
import calendar
from collections import defaultdict
from typing import Dict, List, Any, Tuple
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Auditors mapping according to QA FECT rules
AUDITOR_MAPPING = {
    "VP FE": "ThaoTDP",
    "FSW": "ThaoTDP",
    "FSB": "ThaoTDP",
    "FGW": "DiCQ",
    "FSC CT": "ThanhNTD6",
    "FSC ST": "ThaoTDP + DiCQ",
    "FSC HG": "DiCQ + ThanhNTD6",
}


def get_individual_auditors(auditor_str: str) -> List[str]:
    """Extracts individual auditor names from single or composite strings (e.g. 'ThaoTDP + DiCQ')."""
    parts = auditor_str.replace("+", ",").split(",")
    return [p.strip() for p in parts if p.strip()]


def get_auditor(unit: str, department: str = "") -> str:
    """
    Returns the designated QA auditor for a unit and department based on QA FECT rules:
    - ThaoTDP (Trần Đoàn Phương Thảo):
      + VPFE (NS; Văn Thư; VHDT)
      + FSW
      + FPTU: Đào tạo + Khảo thí
      + FSB
    - ThanhNTD6 (Nguyễn Thị Đan Thanh):
      + FSC CT
      + FPTU (các bộ phận còn lại: HTQT, Tuyển sinh, Hành chính, CTSV, DVSV, CNTT, PTUD, Thư viện, Tài sản, v.v.)
    - DiCQ (Chung Quốc Di):
      + FGW
      + FPTU: 03 Bộ môn
    - ThaoTDP + DiCQ:
      + FSC ST
    - DiCQ + ThanhNTD6:
      + FSC HG
    """
    u = unit.strip().upper()
    dept = department.strip().lower()

    if "VP FE" in u or "VPFE" in u:
        return "ThaoTDP"
    if "FSW" in u or "SWINBURNE" in u:
        return "ThaoTDP"
    if "FSB" in u:
        return "ThaoTDP"
    if "FGW" in u:
        return "DiCQ"
    if "FSC" in u:
        if "ST" in u:
            return "ThaoTDP + DiCQ"
        if "HG" in u:
            return "DiCQ + ThanhNTD6"
        return "ThanhNTD6"
    if "FPTU" in u:
        # DiCQ: 03 Bộ môn
        if "bộ môn" in dept or "bo mon" in dept or "bm" in dept.split():
            return "DiCQ"
        # ThaoTDP: Đào tạo + Khảo thí
        training_keywords = ["đào tạo", "dao tao", "khảo thí", "khao thi", "tc&qlđt", "tc&qldt", "qlđt", "qldt"]
        if any(kw in dept for kw in training_keywords):
            return "ThaoTDP"
        # ThanhNTD6: các bộ phận còn lại của FPTU
        return "ThanhNTD6"

    return "ThaoTDP"


def get_valid_working_days(year: int, month: int) -> List[datetime.date]:
    """
    Returns working days (Mon-Fri) strictly between day 6 and day 23 inclusive.
    Excludes days 1-5, days 24-31, and Saturdays/Sundays.
    """
    valid = []
    for day in range(6, 24):
        try:
            d = datetime.date(year, month, day)
            if d.weekday() < 5:  # Mon=0..Fri=4
                valid.append(d)
        except ValueError:
            break
    return valid


def schedule_tasks(
    tasks_by_unit: Dict[str, List[Dict[str, Any]]],
    year: int = 2026,
    month: int = 8
) -> Tuple[Dict[str, List[Dict[str, Any]]], List[Dict[str, Any]]]:
    """
    Schedules tasks into slots from day 6 to day 23:
    1. Only Mon-Fri.
    2. Strictly day 6 to 23.
    3. Groups by department/auditor: <= 4 tasks per session (half-day), 5-8 tasks per full day.
    4. Ensures slots in calendar never mix different auditors:
       - Each slot is dedicated to a single auditor.
       - ThaoTDP -> White (FFFFFFFF)
       - DiCQ -> Light yellow (FFF7FFB9)
       - ThanhNTD6 -> Light brown (FFFEF4EC)
    Returns:
      (tasks_by_unit, calendar_events)
    """
    valid_days = get_valid_working_days(year, month)
    if not valid_days:
        return tasks_by_unit, []

    # 1. Create work packages per (unit, auditor)
    # Full-day packages: 5 to 8 tasks (occupies both Morning and Afternoon of a day, time '09h00 - 17h00')
    # Half-day packages: 1 to 4 tasks (occupies Sáng '09h00 - 12h00' or Chiều '13h00 - 17h30')
    packages = []
    for unit, task_list in tasks_by_unit.items():
        dept_groups = defaultdict(list)
        for t in task_list:
            dept = t.get("department", "").strip()
            dept_groups[dept].append(t)

        remaining_by_auditor = defaultdict(list)
        for dept, items in dept_groups.items():
            aud = get_auditor(unit, dept)
            rem = list(items)
            # Departments with >= 5 tasks get full-day packages (up to 8 tasks per day)
            while len(rem) >= 5:
                take_cnt = min(8, len(rem))
                packages.append({
                    "type": "full_day",
                    "unit": unit,
                    "department": dept,
                    "auditor": aud,
                    "items": rem[:take_cnt],
                })
                rem = rem[take_cnt:]
            if rem:
                # If 3 or 4 items, keep as individual half-day package for this dept
                if len(rem) >= 3:
                    packages.append({
                        "type": "half_day",
                        "unit": unit,
                        "department": dept,
                        "auditor": aud,
                        "items": rem,
                    })
                else:
                    remaining_by_auditor[aud].extend(rem)

        # Remaining 1-2 items grouped together under same unit & auditor (up to 4 items per session)
        for aud, s_items in remaining_by_auditor.items():
            rem = list(s_items)
            while rem:
                take_cnt = min(4, len(rem))
                chunk = rem[:take_cnt]
                rem = rem[take_cnt:]
                depts = list(dict.fromkeys(t.get("department", "").strip() for t in chunk if t.get("department", "").strip()))
                dept_lbl = "; ".join(depts[:2]) if depts else ""
                packages.append({
                    "type": "half_day",
                    "unit": unit,
                    "department": dept_lbl,
                    "auditor": aud,
                    "items": chunk,
                })

    # 2. Separate full_day and half_day packages
    full_day_pkgs = [p for p in packages if p["type"] == "full_day"]
    half_day_pkgs = [p for p in packages if p["type"] == "half_day"]

    # Slot tracking: slot_assignments[(day_idx, session)] = package
    # Crucial rule: A slot cannot be shared by different auditors!
    slot_assignments = {}
    day_has_full_day = set()
    unit_busy = set()  # (day_idx, session, unit)

    # 2a. Schedule full-day packages onto dedicated days
    for pkg in full_day_pkgs:
        unit = pkg["unit"]
        auditor = pkg["auditor"]
        assigned_day_idx = None
        for day_idx in range(len(valid_days)):
            if day_idx in day_has_full_day:
                continue
            if (day_idx, "Sáng") in slot_assignments or (day_idx, "Chiều") in slot_assignments:
                continue
            if (day_idx, "Sáng", unit) in unit_busy or (day_idx, "Chiều", unit) in unit_busy:
                continue
            assigned_day_idx = day_idx
            break

        if assigned_day_idx is not None:
            day_has_full_day.add(assigned_day_idx)
            slot_assignments[(assigned_day_idx, "Sáng")] = pkg
            slot_assignments[(assigned_day_idx, "Chiều")] = pkg
            unit_busy.add((assigned_day_idx, "Sáng", unit))
            unit_busy.add((assigned_day_idx, "Chiều", unit))
        else:
            # If no free whole day, split into two half-day packages
            half_day_pkgs.append({
                "type": "half_day",
                "unit": pkg["unit"],
                "department": pkg["department"],
                "auditor": pkg["auditor"],
                "items": pkg["items"][:4]
            })
            if len(pkg["items"]) > 4:
                half_day_pkgs.append({
                    "type": "half_day",
                    "unit": pkg["unit"],
                    "department": pkg["department"],
                    "auditor": pkg["auditor"],
                    "items": pkg["items"][4:]
                })

    # 2b. Schedule half-day packages
    # Strict single-auditor rule: only assign to an empty slot or a slot with the EXACT same auditor
    for pkg in half_day_pkgs:
        unit = pkg["unit"]
        auditor = pkg["auditor"]
        assigned = None

        # Priority 1: Completely empty slot with no unit conflict
        for day_idx in range(len(valid_days)):
            if day_idx in day_has_full_day:
                continue
            for session in ["Sáng", "Chiều"]:
                if (day_idx, session) not in slot_assignments:
                    if (day_idx, session, unit) not in unit_busy:
                        assigned = (day_idx, session)
                        break
            if assigned:
                break

        # Priority 2: Empty slot allowing same unit if needed
        if not assigned:
            for day_idx in range(len(valid_days)):
                if day_idx in day_has_full_day:
                    continue
                for session in ["Sáng", "Chiều"]:
                    if (day_idx, session) not in slot_assignments:
                        assigned = (day_idx, session)
                        break
                if assigned:
                    break

        # Priority 3: Merge with existing slot of the SAME auditor (never different auditor)
        if not assigned:
            for (day_idx, session), ex_pkg in list(slot_assignments.items()):
                if day_idx not in day_has_full_day and ex_pkg["auditor"] == auditor:
                    ex_pkg["items"].extend(pkg["items"])
                    assigned = (day_idx, session)
                    break

        if assigned and assigned not in slot_assignments:
            slot_assignments[assigned] = pkg
            unit_busy.add((assigned[0], assigned[1], unit))

    # 3. Build scheduled tasks and calendar events
    calendar_events = []
    processed_pkgs = set()

    for (day_idx, session), pkg in slot_assignments.items():
        pkg_id = id(pkg)
        if pkg_id in processed_pkgs:
            continue
        processed_pkgs.add(pkg_id)

        d_obj = valid_days[day_idx]
        day_str = f"{d_obj.day:02d}.{d_obj.month}.{d_obj.year}"
        is_full = day_idx in day_has_full_day and pkg["type"] == "full_day"

        if is_full:
            time_str = "09h00 - 17h00"
            sess_str = "Cả ngày"
        elif session == "Sáng":
            time_str = "09h00 - 12h00"
            sess_str = "Sáng"
        else:
            time_str = "13h00 - 17h30"
            sess_str = "Chiều"

        items = pkg["items"]
        for item in items:
            item["scheduled_date"] = day_str
            item["scheduled_time"] = time_str
            item["scheduled_auditor"] = pkg["auditor"]
            item["scheduled_session"] = sess_str
            item["scheduled_date_obj"] = d_obj

        docs = []
        for it in items:
            d_val = it.get("document", "").strip()
            if d_val and d_val not in docs:
                docs.append(d_val)
        contents = []
        for it in items:
            c_val = it.get("content", "").strip()
            if c_val and c_val not in contents:
                contents.append(c_val)

        calendar_events.append({
            "date": day_str,
            "date_obj": d_obj,
            "session": sess_str,
            "time": time_str,
            "unit": pkg["unit"],
            "department": pkg["department"],
            "auditor": pkg["auditor"],
            "documents": docs,
            "contents": contents,
            "task_count": len(items)
        })

    return tasks_by_unit, calendar_events


# Color definitions according to QA FECT rules:
# - Màu trắng cho ThaoTDP (0xFFFFFFFF)
# - Màu vàng nhạt cho DiCQ (0xFFF7FFB9)
# - Màu nâu nhạt cho ThanhNTD6 (0xFFFFEF4EC)
FILL_WHITE = PatternFill(fill_type="solid", start_color="FFFFFFFF", end_color="FFFFFFFF")
FILL_YELLOW = PatternFill(fill_type="solid", start_color="FFF7FFB9", end_color="FFF7FFB9")
FILL_BROWN = PatternFill(fill_type="solid", start_color="FFFEF4EC", end_color="FFFEF4EC")

THIN_BORDER_SIDE = Side(border_style="thin", color="B0C4DE")
CAL_CELL_BORDER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)


def get_slot_fill(auditor_str: str) -> PatternFill:
    aud = re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", auditor_str.strip())
    # Composite auditor teams
    if "ThaoTDP + DiCQ" in aud:
        return FILL_WHITE
    if "DiCQ + ThanhNTD6" in aud:
        return FILL_BROWN
    # Single auditor
    if "DiCQ" in aud and "ThaoTDP" not in aud and "ThanhNTD6" not in aud:
        return FILL_YELLOW
    if "ThanhNTD6" in aud and "DiCQ" not in aud and "ThaoTDP" not in aud:
        return FILL_BROWN
    if "ThaoTDP" in aud and "DiCQ" not in aud and "ThanhNTD6" not in aud:
        return FILL_WHITE
    # Fallback checks
    if "DiCQ" in aud:
        return FILL_YELLOW
    if "ThanhNTD6" in aud:
        return FILL_BROWN
    return FILL_WHITE


def populate_calendar_grid(ws, calendar_events: List[Dict[str, Any]], year: int, month_num: int):
    """
    Populates calendar events into the monthly calendar sheet according to user requirements:
    1. Clear background of all cells in calendar, assign fill color:
       - White for ThaoTDP
       - Light yellow (#F7FFB9) for DiCQ
       - Light brown (#FEF4EC) for ThanhNTD6
    2. In calendar with 2 rows S and C:
       - If time is 09h00 - 17h00 (full day): merge S and C cells of that day and paste data
       - If S (Morning): time is 09h00 - 12h00
       - If C (Afternoon): time is 13h00 - 17h30
    3. Clear all data in Columns I, J, K of this sheet.
    """
    # 3. Clear all data in Columns I, J, K (and columns >= 9)
    col_ijk_merges = [
        m for m in list(ws.merged_cells.ranges)
        if m.max_col >= 9
    ]
    for m in col_ijk_merges:
        try:
            ws.unmerge_cells(str(m))
        except Exception:
            pass

    max_r = max(ws.max_row, 35)
    max_c = max(ws.max_column, 15)
    for r in range(1, max_r + 1):
        for c in range(9, max_c + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)

    # Group working days into weeks
    cal = calendar.Calendar(firstweekday=calendar.MONDAY)
    month_days = cal.monthdayscalendar(year, month_num)
    working_weeks = [week for week in month_days if any(d > 0 for d in week[:5])]

    week_row_map = {}
    for w_idx, week in enumerate(working_weeks[:5]):
        d_row = 4 + (w_idx * 3)
        m_row = 5 + (w_idx * 3)
        a_row = 6 + (w_idx * 3)
        for day_idx in range(5):  # Mon=0..Fri=4
            day_val = week[day_idx]
            if day_val > 0:
                col_letter = get_column_letter(2 + day_idx)
                week_row_map[day_val] = {
                    "col": col_letter,
                    "col_idx": 2 + day_idx,
                    "morning_row": m_row,
                    "afternoon_row": a_row,
                }

    cal_font = Font(name="Arial", size=8, bold=False)
    cal_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Clear existing event merges in the calendar grid (Cols B to F, rows 4 to 18)
    grid_merges = [
        m for m in list(ws.merged_cells.ranges)
        if m.min_col >= 2 and m.max_col <= 6 and m.min_row >= 4 and m.max_row <= 18
    ]
    for m in grid_merges:
        try:
            ws.unmerge_cells(str(m))
        except Exception:
            pass

    # 1. Clear values and reset background for all grid cells (item 1: "xoá nền hết các ô trong calendar")
    for w_idx in range(5):
        m_r = 5 + (w_idx * 3)
        a_r = 6 + (w_idx * 3)
        for c_idx in range(2, 7):
            for r in (m_r, a_r):
                cell = ws.cell(row=r, column=c_idx)
                cell.value = None
                cell.fill = PatternFill(fill_type=None)
                cell.border = CAL_CELL_BORDER

    # Group events by day_val
    events_by_day = defaultdict(lambda: defaultdict(list))
    for ev in calendar_events:
        day_num = ev["date_obj"].day
        sess = ev["session"]
        events_by_day[day_num][sess].append(ev)

    def format_event_text(ev_list):
        blocks = []
        for ev in ev_list:
            unit_header = f"{ev['unit']}- {ev['department'] or 'ĐƠN VỊ'} ({ev['time']})- {ev['auditor']}"
            lines = [unit_header]
            if ev["documents"]:
                doc_summary = " & ".join([d.split("\n")[0][:45] for d in ev["documents"][:2]])
                lines.append(f"  * {doc_summary}")
            if ev["contents"]:
                for idx, c in enumerate(ev["contents"][:2], 1):
                    c_clean = c.split("\n")[0][:40]
                    lines.append(f"  {idx}. {c_clean}")
            blocks.append("\n".join(lines))
        return "\n\n".join(blocks)

    for day_num, sess_map in events_by_day.items():
        pos = week_row_map.get(day_num)
        if not pos:
            continue

        c_idx = pos["col_idx"]
        m_row = pos["morning_row"]
        a_row = pos["afternoon_row"]

        # Check if this day has a full-day event ("Cả ngày" or "09h00 - 17h00")
        has_full_day = "Cả ngày" in sess_map

        if has_full_day:
            # 2. Merge morning and afternoon cells for this day and paste data
            all_day_evs = sess_map["Cả ngày"] + sess_map.get("Sáng", []) + sess_map.get("Chiều", [])
            text = format_event_text(all_day_evs)
            auditor_str = ", ".join(e["auditor"] for e in all_day_evs)
            fill_to_apply = get_slot_fill(auditor_str)

            ws.merge_cells(start_row=m_row, start_column=c_idx, end_row=a_row, end_column=c_idx)

            top_cell = ws.cell(row=m_row, column=c_idx)
            top_cell.value = text
            top_cell.font = cal_font
            top_cell.alignment = cal_align
            top_cell.fill = fill_to_apply
            top_cell.border = CAL_CELL_BORDER

            bottom_cell = ws.cell(row=a_row, column=c_idx)
            bottom_cell.fill = fill_to_apply
            bottom_cell.border = CAL_CELL_BORDER
        else:
            # Separate Sáng (09h00 - 12h00) and Chiều (13h00 - 17h30)
            if "Sáng" in sess_map:
                s_evs = sess_map["Sáng"]
                s_text = format_event_text(s_evs)
                s_aud = ", ".join(e["auditor"] for e in s_evs)
                s_fill = get_slot_fill(s_aud)

                s_cell = ws.cell(row=m_row, column=c_idx)
                s_cell.value = s_text
                s_cell.font = cal_font
                s_cell.alignment = cal_align
                s_cell.fill = s_fill
                s_cell.border = CAL_CELL_BORDER

            if "Chiều" in sess_map:
                c_evs = sess_map["Chiều"]
                c_text = format_event_text(c_evs)
                c_aud = ", ".join(e["auditor"] for e in c_evs)
                c_fill = get_slot_fill(c_aud)

                c_cell = ws.cell(row=a_row, column=c_idx)
                c_cell.value = c_text
                c_cell.font = cal_font
                c_cell.alignment = cal_align
                c_cell.fill = c_fill
                c_cell.border = CAL_CELL_BORDER
