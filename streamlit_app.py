import io
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from grader import grade_file

st.set_page_config(
    page_title="Football Excel Exam Grader",
    page_icon="⚽",
    layout="wide",
)

st.title("⚽ Football Excel Exam Auto-Grader")
st.caption("Upload one or more student .xlsx files – the tool grades them and returns colour-coded corrected files.")

with st.sidebar:
    st.header("⚙️ Options")
    do_recalc = st.checkbox(
        "Re-calculate formulas with LibreOffice",
        value=True,
        help="Recommended. Ensures student formulas are evaluated before comparison.",
    )
    st.markdown("---")
    st.markdown("**Colour legend in output file:**")
    st.markdown("🟩 Green = correct (≥80%)")
    st.markdown("🟨 Yellow = partial (40–79%) or manual review")
    st.markdown("🟥 Red = incorrect (<40%)")
    st.markdown("---")
    st.markdown("**Auto-graded:** S1 Q0–Q10 · S2 Q1–Q6, Q8–Q14 · S3 Q1–Q10 · S4 Q1–Q2, Q6")
    st.markdown("**Manual review:** S2 Q7 · S4 Q3–Q5, Q7–Q8")

uploaded = st.file_uploader(
    "Drop student Excel files here",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded:
    st.info("👆 Upload at least one student .xlsx file to start grading.")
    st.stop()

if st.button("🚀 Grade all files", type="primary"):

    graded_files = []
    results_all = {}
    progress = st.progress(0, text="Grading…")

    for idx, up in enumerate(uploaded):
        progress.progress(idx / len(uploaded), text=f"Grading {up.name}…")
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(up.read())
            tmp_path = Path(tmp.name)
        try:
            res = grade_file(tmp_path, recalc=do_recalc)
            results_all[up.name] = res
            graded_files.append((up.name, res["output_path"]))
        except Exception as e:
            st.error(f"Error grading **{up.name}**: {e}")

    progress.progress(1.0, text="Done ✓")

    if not results_all:
        st.stop()

    # ── summary table ────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📊 Grading Results")

    rows = []
    for fname, res in results_all.items():
        row = {"File": fname}
        grand_earned = grand_max_auto = grand_max_total = 0
        for sec, data in res["summary"].items():
            pct = data["score"] / data["max_auto"] * 100 if data["max_auto"] else 0
            row[sec] = f"{pct:.0f}%  ({data['score']:.1f}/{data['max_auto']})"
            grand_earned    += data["score"]
            grand_max_auto  += data["max_auto"]
            grand_max_total += data["max_total"]
        total_pct = grand_earned / grand_max_auto * 100 if grand_max_auto else 0
        row["TOTAL %"] = f"{total_pct:.1f}%  ({grand_earned:.1f}/{grand_max_auto} auto)"
        rows.append(row)

    st.dataframe(pd.DataFrame(rows), use_container_width=True)

    # ── per-file breakdown ───────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🔍 Question-by-question breakdown")

    for fname, res in results_all.items():
        with st.expander(f"📄 {fname}"):
            for sec, data in res["summary"].items():
                pct = data["score"] / data["max_auto"] * 100 if data["max_auto"] else 0
                st.markdown(
                    f"**{sec}** – {pct:.0f}%  "
                    f"({data['score']:.1f}/{data['max_auto']} auto-graded pts | "
                    f"{data['max_total']} total pts)"
                )
                q_rows = []
                for q, v in data["questions"].items():
                    if v["rate"] is None:
                        status = "⚠️"
                        pts = "–"
                    elif v["rate"] >= 0.8:
                        status = "✅"
                        pts = f"{v['max'] * v['rate']:.1f}/{v['max']}"
                    elif v["rate"] >= 0.4:
                        status = "🟡"
                        pts = f"{v['max'] * v['rate']:.1f}/{v['max']}"
                    else:
                        status = "❌"
                        pts = f"0/{v['max']}"
                    q_rows.append({
                        "Q": q,
                        "Max": v["max"],
                        "Pts": pts,
                        "Status": status,
                        "Detail": v["detail"],
                    })
                st.dataframe(pd.DataFrame(q_rows), use_container_width=True, hide_index=True)
                if data["manual_questions"]:
                    st.warning(f"⚠️ Manual review needed for: {', '.join(data['manual_questions'])}")

    # ── downloads ────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("⬇️ Download Graded Files")

    if len(graded_files) == 1:
        fname, path = graded_files[0]
        with open(path, "rb") as f:
            st.download_button(
                label=f"Download GRADED_{fname}",
                data=f.read(),
                file_name=f"GRADED_{fname}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for fname, path in graded_files:
                zf.write(path, arcname=f"GRADED_{fname}")
        zip_buf.seek(0)
        st.download_button(
            label=f"📦 Download all {len(graded_files)} graded files (.zip)",
            data=zip_buf.read(),
            file_name="graded_exams.zip",
            mime="application/zip",
            type="primary",
        )
        cols = st.columns(min(4, len(graded_files)))
        for i, (fname, path) in enumerate(graded_files):
            with cols[i % len(cols)]:
                with open(path, "rb") as f:
                    st.download_button(
                        label=fname[:25],
                        data=f.read(),
                        file_name=f"GRADED_{fname}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
