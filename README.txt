# SEO Content Mapper — Setup Guide

## What this tool does
Maps untapped keywords (from Semrush keyword gap analysis) to your existing 
pages based on URL slug, Page Title, Meta Description and H1.
Produces a formatted Excel file with mapped URLs, match scores and a 
content gap analysis.

---

## One-time setup (5 minutes)

### Step 1 — Install Python
If you don't have Python installed:
Download from https://python.org/downloads (get version 3.10 or higher)
During install, tick "Add Python to PATH"

### Step 2 — Open Terminal / Command Prompt
Windows: Press Win+R, type cmd, press Enter
Mac: Press Cmd+Space, type Terminal, press Enter

### Step 3 — Install required packages
Copy and paste this command, then press Enter:

    pip install streamlit pandas numpy scikit-learn openpyxl

Wait for it to finish (1-2 minutes).

### Step 4 — Run the app
Navigate to the folder containing app.py:

    cd path/to/seo_mapper

Then run:

    streamlit run app.py

Your browser will open automatically at http://localhost:8501

---

## Every time you use it

1. Open Terminal / Command Prompt
2. Run: streamlit run app.py
3. Browser opens — upload your files and click Run Mapping
4. Download the Excel output

---

## Input files needed

### File 1 — GSC Data (Excel .xlsx)
Export from Google Search Console / Looker Studio
Required columns:
- Query
- Landing Page
- Url Clicks (or Clicks)
- Impressions
- Page Title
- Meta Description
- H1

### File 2 — Keywords + URLs (Excel .xlsx)
Two sheets in one file:

Sheet 1 (keywords):
- Keyword
- Volume (or Search Volume)

Sheet 2 (URL index):
- Landing Page
- Page Title
- Meta Description
- H1

---

## Settings (adjustable in the sidebar)

Match Score Threshold: 0.15 (default)
- Lower = more matches but weaker quality
- Higher = fewer matches but stronger quality
- Recommended range: 0.15 to 0.25

Content Weights:
- URL Slug: 5 (default) — how much the URL path influences matching
- Page Title: 3 (default)
- H1: 2 (default)
- Meta Description: 1 (default)

---

## Output file

Tab 1 — Keyword Mapping:
Columns: Keyword, Search Volume, Landing Page, Intent, Match Source, Match Score
Sorted by Match Score descending
Colours: Green=service page, Blue=blog, Yellow=GSC fallback, White=unmapped

Tab 2 — Content Gap Analysis:
Unmapped keywords grouped by topic with total search volume
Priority flags: High (>200k vol), Medium (>50k), Low (<50k)
Page type needed: Service Page or Blog/Guide

---

## Score guide

0.50+       Strong — fully trust
0.30-0.50   Good — trust it
0.20-0.30   Acceptable — spot check
0.15-0.20   Weak — manually verify
0.40 (GSC)  Fallback — review before using
Blank       Content gap — no existing page matches

---

## Troubleshooting

"streamlit not found" — run: pip install streamlit
"No module named pandas" — run: pip install pandas numpy scikit-learn openpyxl
Browser doesn't open — go to http://localhost:8501 manually
Column not found — use the column mapping dropdowns in the tool to select your columns

---

## Questions
All the logic in this tool was built and validated across a full SEO 
content mapping workflow. The taxonomy rules, conflict prevention, 
intent classification and scoring system are all baked in.
