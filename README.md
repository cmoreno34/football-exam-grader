# ⚽ Football Excel Exam Auto-Grader

Automatically grades student Excel exam submissions and returns colour-coded corrected files.

## 🚀 Deployed on Streamlit Cloud

Upload student `.xlsx` files → get scores + graded files back instantly.

## 📁 Repo structure

```
├── app.py              ← Streamlit UI
├── grader.py           ← Grading engine
├── solution.xlsx       ← Reference solution
├── requirements.txt    ← Python deps
├── packages.txt        ← System deps (LibreOffice)
└── scripts/
    ├── recalc.py
    └── office/
        └── soffice.py
```

## ⚙️ Deploy on Streamlit Cloud

1. Fork / upload this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. New app → select this repo → `app.py` → Deploy
