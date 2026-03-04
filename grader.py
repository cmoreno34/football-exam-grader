"""
Football Excel Exam Auto-Grader
Compares student submissions against the solution file and fills in grading cells.
"""

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ── paths ──────────────────────────────────────────────────────────────────
SOLUTION_PATH = Path(__file__).parent / "solution.xlsx"
SCRIPTS_DIR   = Path(__file__).parent / "scripts"

# ── colour helpers ──────────────────────────────────────────────────────────
GREEN  = PatternFill("solid", start_color="C6EFCE")
RED    = PatternFill("solid", start_color="FFC7CE")
YELLOW = PatternFill("solid", start_color="FFEB9C")
GREEN_FONT  = Font(color="276221", bold=True)
RED_FONT    = Font(color="9C0006", bold=True)
YELLOW_FONT = Font(color="9C5700", bold=True)

# ── tolerance for numeric comparison ───────────────────────────────────────
NUM_TOL = 0.02   # 2 % relative tolerance


def _eq_val(a, b):
    """Return True if two cell values are considered equal."""
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    try:
        fa, fb = float(a), float(b)
        if fb == 0:
            return abs(fa) < 1e-6
        return abs(fa - fb) / abs(fb) <= NUM_TOL
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def _col_values(ws, col, row_start, row_end):
    return [ws.cell(r, col).value for r in range(row_start, row_end + 1)]


def _rect_values(ws, row_start, row_end, col_start, col_end):
    return [
        [ws.cell(r, c).value for c in range(col_start, col_end + 1)]
        for r in range(row_start, row_end + 1)
    ]


def _match_rate(student_vals, sol_vals):
    """Fraction of cells that match (flattened lists / nested lists)."""
    if not isinstance(student_vals[0], list):
        pairs = list(zip(student_vals, sol_vals))
    else:
        flat_s = [v for row in student_vals for v in row]
        flat_r = [v for row in sol_vals   for v in row]
        pairs  = list(zip(flat_s, flat_r))
    if not pairs:
        return 0.0
    return sum(_eq_val(a, b) for a, b in pairs) / len(pairs)


def recalculate(path: Path) -> Path:
    """Run LibreOffice recalc on a copy of the file; return path to recalculated copy."""
    tmp = Path(tempfile.mkdtemp()) / path.name
    shutil.copy(path, tmp)
    try:
        result = subprocess.run(
            ["python3", str(SCRIPTS_DIR / "recalc.py"), str(tmp)],
            capture_output=True, text=True, timeout=60,
            cwd=str(SCRIPTS_DIR.parent)
        )
        if result.returncode != 0:
            print(f"[WARN] recalc returned {result.returncode}: {result.stderr[:200]}")
    except Exception as e:
        print(f"[WARN] recalc failed: {e}")
    return tmp


# ── grading specification ──────────────────────────────────────────────────
# Each entry:
#   sheet, score_col, score_row, label, max_pts, check_fn(student_ws, sol_ws) -> (0|1, detail)

def _check_named_ranges(student_wb, expected_names):
    """Return (1, detail) if all expected named ranges exist."""
    defined = {n.lower() for n in student_wb.defined_names.keys()}
    missing = [n for n in expected_names if n.lower() not in defined]
    if not missing:
        return 1, "✔ All named ranges found"
    return 0, f"✘ Missing: {', '.join(missing)}"


def _check_table_exists(student_wb, table_name):
    for ws in student_wb.worksheets:
        for tbl in ws.tables.values():
            if tbl.name.lower() == table_name.lower():
                return 1, f"✔ Table '{table_name}' found"
    return 0, f"✘ Table '{table_name}' not found"


def grade_file(student_path: Path, recalc: bool = True) -> dict:
    """
    Grade a student file against the solution.
    Returns a dict with per-question results and writes scores into a copy
    of the student file (returned as a Path).
    """
    # ── recalculate student file ─────────────────────────────────────────
    if recalc:
        calc_path = recalculate(student_path)
    else:
        calc_path = student_path

    sol_path  = recalculate(SOLUTION_PATH) if recalc else SOLUTION_PATH

    sol_wb_d  = load_workbook(str(sol_path),  data_only=True)
    stu_wb_d  = load_workbook(str(calc_path), data_only=True)
    stu_wb    = load_workbook(str(calc_path))   # write scores onto recalculated copy

    results = {}   # section -> {q_label -> {score, max, detail}}

    # ════════════════════════════════════════════════════════════════════
    # SECTION 1
    # ════════════════════════════════════════════════════════════════════
    s1_sol = sol_wb_d["Section 1 "]
    s1_stu = stu_wb_d["Section 1 "] if "Section 1 " in stu_wb_d.sheetnames else None
    s1_out = stu_wb["Section 1 "]    if "Section 1 " in stu_wb.sheetnames   else None

    S1 = {}

    # Q0 – Named ranges
    expected_names = ["PlayerName","Goals","Assists","Matches","Salary",
                      "MarketValue","YellowCards","RedCards","Position"]
    mark0, det0 = _check_named_ranges(stu_wb_d, expected_names)
    S1["Q0"] = {"max": 1, "score": mark0, "detail": det0}

    if s1_stu:
        checks = [
            ("Q1",  1, 12, 16, 27, False),   # Total Points
            ("Q2",  2, 13, 16, 27, True),     # Points/Match (numeric)
            ("Q3",  2, 14, 16, 27, True),     # Discipline Score
            ("Q4",  2, 15, 16, 27, False),    # Performance Rating (text)
            ("Q5",  5, 16, 16, 27, True),     # Value Per Point
            ("Q6",  2, 17, 16, 27, True),     # Cost Per Point
            ("Q7",  5, 18, 16, 27, False),    # Contract Status (text)
            ("Q8",  5, 19, 16, 27, False),    # Investment Priority
            ("Q9",  5, 20, 16, 27, False),    # Star/Regular
            ("Q10", 5, 21, 16, 27, True),     # LAMBDA Value
        ]
        for label, max_pts, col, r0, r1, numeric in checks:
            sv = _col_values(s1_stu, col, r0, r1)
            rv = _col_values(s1_sol, col, r0, r1)
            rate = _match_rate(sv, rv)
            mark = 1 if rate >= 0.8 else 0
            S1[label] = {
                "max": max_pts,
                "score": mark,
                "detail": f"{'✔' if mark else '✘'} {int(rate*100)}% cells correct"
            }
    else:
        for q in ["Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8","Q9","Q10"]:
            S1[q] = {"max": [1,2,2,2,5,2,5,5,5,5][int(q[1:])-1], "score": 0, "detail": "✘ Sheet not found"}

    results["Section 1"] = S1

    # Write scores to Section 1
    if s1_out:
        q_order = ["Q0","Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8","Q9","Q10"]
        for i, q in enumerate(q_order):
            row = 3 + i
            mark = S1[q]["score"]
            cell = s1_out.cell(row, 25)   # column Y
            cell.value = mark
            cell.fill  = GREEN if mark else RED
            cell.font  = GREEN_FONT if mark else RED_FONT

    # ════════════════════════════════════════════════════════════════════
    # SECTION 2
    # ════════════════════════════════════════════════════════════════════
    s2_sol = sol_wb_d["Section 2"]
    s2_stu = stu_wb_d["Section 2"] if "Section 2" in stu_wb_d.sheetnames else None
    s2_out = stu_wb["Section 2"]   if "Section 2" in stu_wb.sheetnames   else None

    S2 = {}

    if s2_stu:
        # Q1 – Team/League/Stadium (cols C-E, rows 15-22)
        sv = _rect_values(s2_stu, 15, 22, 3, 5)
        rv = _rect_values(s2_sol, 15, 22, 3, 5)
        r  = _match_rate(sv, rv)
        S2["Q1"] = {"max":1, "score": 1 if r>=0.8 else 0, "detail": f"{'✔' if r>=0.8 else '✘'} {int(r*100)}% cells"}

        # Q2 – Capacity (col F)
        sv = _col_values(s2_stu, 6, 15, 22)
        rv = _col_values(s2_sol, 6, 15, 22)
        r  = _match_rate(sv, rv)
        S2["Q2"] = {"max":2, "score": 1 if r>=0.8 else 0, "detail": f"{'✔' if r>=0.8 else '✘'} {int(r*100)}% cells"}

        # Q3 – Ticket Price (col G)
        sv = _col_values(s2_stu, 7, 15, 22)
        rv = _col_values(s2_sol, 7, 15, 22)
        r  = _match_rate(sv, rv)
        S2["Q3"] = {"max":2, "score": 1 if r>=0.8 else 0, "detail": f"{'✔' if r>=0.8 else '✘'} {int(r*100)}% cells"}

        # Q4 – Coach + Revenue (cols H-I)
        sv = _rect_values(s2_stu, 15, 22, 8, 9)
        rv = _rect_values(s2_sol, 15, 22, 8, 9)
        r  = _match_rate(sv, rv)
        S2["Q4"] = {"max":1, "score": 1 if r>=0.8 else 0, "detail": f"{'✔' if r>=0.8 else '✘'} {int(r*100)}% cells"}

        # Q5 – Attendance % (col J)
        sv = _col_values(s2_stu, 10, 15, 22)
        rv = _col_values(s2_sol, 10, 15, 22)
        r  = _match_rate(sv, rv)
        S2["Q5"] = {"max":1, "score": 1 if r>=0.8 else 0, "detail": f"{'✔' if r>=0.8 else '✘'} {int(r*100)}% cells"}

        # Q6 – Drop-downs: check data validation exists in B26/D26
        try:
            has_dv = len(list(s2_stu.data_validations.dataValidation)) >= 2
            S2["Q6"] = {"max":1, "score": 1 if has_dv else 0,
                        "detail": "✔ Drop-downs detected" if has_dv else "✘ Drop-downs missing – MANUAL REVIEW"}
        except Exception:
            S2["Q6"] = {"max":1, "score": 0, "detail": "⚠ Could not check – MANUAL REVIEW"}

        # Q7 – Conditional formatting (manual)
        S2["Q7"] = {"max":5, "score": 0, "detail": "⚠ MANUAL REVIEW – conditional formatting"}

        # Q8 – FILTER+CHOOSECOLS: compare B32:C39 region
        sv_q8 = _rect_values(s2_stu, 32, 39, 2, 3)
        rv_q8 = _rect_values(s2_sol, 32, 39, 2, 3)
        r8 = _match_rate(sv_q8, rv_q8)
        S2["Q8"] = {"max":5, "score": 1 if r8>=0.7 else 0, "detail": f"{'✔' if r8>=0.7 else '✘'} {int(r8*100)}% dynamic array cells"}

        # Q9 – LET+FILTER: compare E32:F35 region
        sv_q9 = _rect_values(s2_stu, 32, 35, 5, 6)
        rv_q9 = _rect_values(s2_sol, 32, 35, 5, 6)
        r9 = _match_rate(sv_q9, rv_q9)
        S2["Q9"] = {"max":5, "score": 1 if r9>=0.7 else 0, "detail": f"{'✔' if r9>=0.7 else '✘'} {int(r9*100)}% dynamic array cells"}

        # Q10 – VSTACK: compare H32:I35 region
        sv_q10 = _rect_values(s2_stu, 32, 35, 8, 9)
        rv_q10 = _rect_values(s2_sol, 32, 35, 8, 9)
        r10 = _match_rate(sv_q10, rv_q10)
        S2["Q10"] = {"max":5, "score": 1 if r10>=0.7 else 0, "detail": f"{'✔' if r10>=0.7 else '✘'} {int(r10*100)}% dynamic array cells"}

        # Q11-Q14 – Multi-criteria, answers in col E rows 41-44
        q_extra = [("Q11",2,47), ("Q12",5,48), ("Q13",5,49), ("Q14",5,50)]
        for label, max_pts, row in q_extra:
            sv_v = s2_stu.cell(row, 5).value
            rv_v = s2_sol.cell(row, 5).value
            mark = 1 if _eq_val(sv_v, rv_v) else 0
            S2[label] = {"max": max_pts, "score": mark,
                         "detail": f"{'✔' if mark else '✘'} Student={sv_v} | Expected={rv_v}"}
    else:
        for q, m in [("Q1",1),("Q2",2),("Q3",2),("Q4",1),("Q5",1),("Q6",1),
                     ("Q7",5),("Q8",5),("Q9",5),("Q10",5),("Q11",2),("Q12",5),("Q13",5),("Q14",5)]:
            S2[q] = {"max": m, "score": 0, "detail": "✘ Sheet not found"}

    results["Section 2"] = S2

    # Write scores to Section 2
    if s2_out:
        for i, q in enumerate(["Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8","Q9","Q10"]):
            row  = 3 + i
            mark = S2[q]["score"]
            cell = s2_out.cell(row, 15)   # column O
            cell.value = mark
            cell.fill  = GREEN if mark else (YELLOW if "MANUAL" in S2[q]["detail"] else RED)
            cell.font  = GREEN_FONT if mark else (YELLOW_FONT if "MANUAL" in S2[q]["detail"] else RED_FONT)
        for i, q in enumerate(["Q11","Q12","Q13","Q14"]):
            row  = 47 + i
            mark = S2[q]["score"]
            cell = s2_out.cell(row, 10)   # column J
            cell.value = mark
            cell.fill  = GREEN if mark else RED
            cell.font  = GREEN_FONT if mark else RED_FONT

    # ════════════════════════════════════════════════════════════════════
    # SECTION 3
    # ════════════════════════════════════════════════════════════════════
    s3_sol = sol_wb_d["Section 3"]
    s3_stu = stu_wb_d["Section 3"] if "Section 3" in stu_wb_d.sheetnames else None
    s3_out = stu_wb["Section 3"]   if "Section 3" in stu_wb.sheetnames   else None

    S3 = {}

    if s3_stu:
        checks3 = [
            ("Q1",  1,  7),   # Clean ID
            ("Q2",  1,  8),   # Team Code
            ("Q3",  2,  9),   # Email
            ("Q4",  2, 10),   # Reg Year
            ("Q5",  2, 11),   # Player Num
            ("Q7",  2, 13),   # Full Name
            ("Q8",  5, 14),   # Username
            ("Q9",  5, 15),   # LET Display
            ("Q10", 5, 16),   # LAMBDA Clean
        ]
        for label, max_pts, col in checks3:
            sv = _col_values(s3_stu, col, 17, 26)
            rv = _col_values(s3_sol, col, 17, 26)
            r  = _match_rate(sv, rv)
            S3[label] = {"max": max_pts, "score": 1 if r>=0.8 else 0,
                         "detail": f"{'✔' if r>=0.8 else '✘'} {int(r*100)}% cells correct"}

        # Q6 – Age: DATEDIF result changes daily; just check formula exists
        ages_stu = _col_values(s3_stu, 12, 17, 26)
        ages_sol = _col_values(s3_sol, 12, 17, 26)
        has_vals = sum(1 for v in ages_stu if v is not None and str(v).isdigit() or isinstance(v, (int,float))) >= 8
        r6 = _match_rate(ages_stu, ages_sol)
        mark6 = 1 if r6 >= 0.7 or has_vals else 0
        S3["Q6"] = {"max": 5, "score": mark6,
                    "detail": f"{'✔' if mark6 else '✘'} Age col: {int(r6*100)}% match (dynamic – based on today's date)"}
    else:
        for q, m in [("Q1",1),("Q2",1),("Q3",2),("Q4",2),("Q5",2),("Q6",5),
                     ("Q7",2),("Q8",5),("Q9",5),("Q10",5)]:
            S3[q] = {"max": m, "score": 0, "detail": "✘ Sheet not found"}

    results["Section 3"] = S3

    # Write scores to Section 3
    if s3_out:
        for i, q in enumerate(["Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8","Q9","Q10"]):
            row  = 28 + i
            mark = S3[q]["score"]
            cell = s3_out.cell(row, 3)   # column C
            cell.value = mark
            cell.fill  = GREEN if mark else RED
            cell.font  = GREEN_FONT if mark else RED_FONT

    # ════════════════════════════════════════════════════════════════════
    # SECTION 4
    # ════════════════════════════════════════════════════════════════════
    s4_sol = sol_wb_d["Section 4"]
    s4_stu = stu_wb_d["Section 4"] if "Section 4" in stu_wb_d.sheetnames else None
    s4_out = stu_wb["Section 4"]   if "Section 4" in stu_wb.sheetnames   else None

    S4 = {}

    # Q1 – Table named MatchResults
    mark_t, det_t = _check_table_exists(stu_wb_d, "MatchResults")
    S4["Q1"] = {"max": 1, "score": mark_t, "detail": det_t}

    if s4_stu:
        # Q2 – GoalDiff + Result Home + Result Away (cols H-J, rows 13-37)
        sv = _rect_values(s4_stu, 13, 37, 8, 10)
        rv = _rect_values(s4_sol, 13, 37, 8, 10)
        r  = _match_rate(sv, rv)
        S4["Q2"] = {"max":1, "score": 1 if r>=0.8 else 0, "detail": f"{'✔' if r>=0.8 else '✘'} {int(r*100)}% cells correct"}

        # Q3 – Color scale CF on Attendance: manual
        S4["Q3"] = {"max":3, "score": 0, "detail": "⚠ MANUAL REVIEW – color scale CF"}

        # Q4 – Icon sets: manual
        S4["Q4"] = {"max":2, "score": 0, "detail": "⚠ MANUAL REVIEW – icon sets CF"}

        # Q5 – Slicer: manual
        S4["Q5"] = {"max":3, "score": 0, "detail": "⚠ MANUAL REVIEW – slicer"}

        # Q6 – AVERAGEIFS in J8
        sv_v = s4_stu.cell(8, 10).value
        rv_v = s4_sol.cell(8, 10).value
        mark6 = 1 if _eq_val(sv_v, rv_v) else 0
        S4["Q6"] = {"max":5, "score": mark6,
                    "detail": f"{'✔' if mark6 else '✘'} Student={sv_v} | Expected={rv_v}"}

        # Q7 – Row highlight CF: manual
        S4["Q7"] = {"max":5, "score": 0, "detail": "⚠ MANUAL REVIEW – row highlight CF"}

        # Q8 – Column highlight CF: manual
        S4["Q8"] = {"max":5, "score": 0, "detail": "⚠ MANUAL REVIEW – column highlight CF"}
    else:
        for q, m in [("Q2",1),("Q3",3),("Q4",2),("Q5",3),("Q6",5),("Q7",5),("Q8",5)]:
            S4[q] = {"max": m, "score": 0, "detail": "✘ Sheet not found"}

    results["Section 4"] = S4

    # Write scores to Section 4
    if s4_out:
        for i, q in enumerate(["Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8"]):
            row  = 14 + i
            mark = S4[q]["score"]
            cell = s4_out.cell(row, 17)   # column Q
            cell.value = mark
            cell.fill  = GREEN if mark else (YELLOW if "MANUAL" in S4[q]["detail"] else RED)
            cell.font  = GREEN_FONT if mark else (YELLOW_FONT if "MANUAL" in S4[q]["detail"] else RED_FONT)

    # ── save output ──────────────────────────────────────────────────────
    out_path = student_path.parent / f"GRADED_{student_path.name}"
    stu_wb.save(str(out_path))

    # ── compute totals ──────────────────────────────────────────────────
    summary = {}
    for sec, data in results.items():
        total_max   = sum(v["max"]   for v in data.values())
        total_score = sum(v["max"] * v["score"] for v in data.values())
        manual_qs   = [q for q, v in data.items() if "MANUAL" in v["detail"]]
        summary[sec] = {
            "score": total_score,
            "max":   total_max,
            "manual_questions": manual_qs,
            "questions": data,
        }

    return {"summary": summary, "output_path": out_path}
