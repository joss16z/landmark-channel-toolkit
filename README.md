# Landmark Channel Toolkit V4

## What's new
- `Floorplate Update Report` supports PowerPoint **and PDF** floorplates.
- Better matching for unit formats like `G.02`, `G02`, `Unit4104`, `Unit 4104`.
- Report sheets: Summary, Price Changed, Missing in Floorplate, Extra in Floorplate, Area Changed, Matched, All Checks.
- Cleaner Excel formatting with colours and frozen headers.
- `PPT Price Update` remains price-only replacement for existing PowerPoint labels.

## Deploy
Upload and overwrite these files in GitHub:
- app.py
- requirements.txt
- README.md

Then commit. Streamlit Cloud will redeploy automatically; URL stays the same.

## Live Floorplate (feature branch; not deployed)

The additive Live Floorplate page includes a source-checked mapping for all 115
Hurstville Lots, five online floor groups and exactly five PowerPoint slides.
Use Source PDF preview with synthetic inventory to review the supplied artwork
without API access. Existing Report Tools remain available. See
[setup, mapping, API configuration and limitations](docs/live-floorplate.md).
Upload the original PDF and reference PPT in the approved runtime. Validation of
the actual getProject response remains pending the redacted sample. Credentials belong only
in server-side Streamlit Secrets in the company-approved environment.
