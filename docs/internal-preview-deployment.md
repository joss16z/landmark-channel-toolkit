# Internal Streamlit test deployment — Draft PR #1

**Branch:** `feature/live-floorplate` (do not merge). **Entrypoint:** `app.py`
in the repository root. **Python:** 3.11 or 3.12. **Live API:** disabled.
No API key, sign, Secrets file or live inventory is needed for this test.

## Simplest option: company-approved Windows computer

Open PowerShell in your existing `landmark-channel-toolkit` folder. Stop if the
first command shows local changes; ask IT to preserve them rather than resetting.

```powershell
git status --short
git fetch origin
git switch feature/live-floorplate
git pull --ff-only origin feature/live-floorplate
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\scripts\Start-InternalPreview.ps1
```

For the initial setup, Python 3.11/3.12 with the Windows `py` launcher must already
be approved/installed. On subsequent starts, run only the last command. If company
PowerShell execution policy blocks the script, ask IT to approve it; do not change
execution policy. The equivalent launch command, without a PowerShell script, is:

```powershell
$env:LANDMARK_PREVIEW_ONLY = '1'
.\.venv\Scripts\python.exe -m streamlit run app.py --server.address=127.0.0.1 --server.port=8501
```

Open your browser at `http://127.0.0.1:8501`. This binds only to your computer and
is not a shared hosted deployment. Keep the PowerShell window running. Press
Ctrl+C there to stop. Do not open a public tunnel or expose the port.

## Upload the supplied originals

1. In the left sidebar, select **Live Floorplate**.
2. Leave **Source PDF preview (synthetic inventory)** selected. This is now the
   default. The live API option is hidden and server-side synchronization is blocked.
3. Under **Original architectural PDF floorplates**, upload the original
   `Hurstville_Floorplate_E-LVL8 - Draft Strata Issue A 27.07.26.pdf` — one file,
   six pages, including its cover. Do not remove the cover or re-save/reorder it.
4. Under **Reference Landmark PowerPoint**, upload
   `Hurstville Rise-Floor plate & Pricing-20261002.pptx`.
5. Leave the optional replacement mapping empty. The checked 115-Lot mapping is
   already included and matched to the original PDF by checksum.
6. Review/check the PDF/mapping confirmation box. The five floor groups and
   synthetic information boxes will appear. No historical PPT price is reused.
7. Select Ground, Level 1, Levels 2–4, Levels 5–7 or Level 8. Use **Refresh synthetic
   source-layout preview** when needed. After five minutes the snapshot expires;
   refresh to show test labels again.
8. Click **Download five-slide as-of PowerPoint**. All test values and TEST units
   are synthetic; this output must not be published as actual availability/prices.

Uploads remain in the active server session; restarting/reconnecting may require
uploading them again. The app does not commit/upload them to GitHub or write them
into the repository. Keep originals and downloaded sample exports in approved
private storage outside Git. The existing Report Tools remain available unchanged.

## Dependencies

Install the root **`requirements.txt`**, including dependencies imported by the
existing Report Tools even when testing only Floorplate preview:

`streamlit>=1.37,<2`, `pandas`, `python-pptx`, `openpyxl`, `xlrd`, `lxml`, `html5lib`,
`beautifulsoup4`, `pdfplumber`, `PyMuPDF`, `pytesseract`, `Pillow`.

The operating-system Tesseract executable is needed only for Report Tools OCR on
image-only price-list PDFs. On Linux the existing **`packages.txt`** specifies
`tesseract-ocr`; on Windows IT must install an approved Tesseract build and add it
to PATH if those OCR flows are to be exercised. Original-source preview itself does
not require OCR. Do not use a missing OCR executable as proof that all Report Tools
have been tested on a new machine.

Optional validation before launch:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -q
```

## Shared internal hosting — IT configuration still required

The owner has selected a separate Streamlit Community Cloud test app. Its actual
private access controls and company hosting approval still need verification. No
hosted app has been created. Follow the [exact Cloud steps](community-cloud-internal-test.md). Do not upload confidential artwork to an
unapproved/public service. A private GitHub repository alone does not authenticate
web viewers or make a Streamlit app private.

For an approved hosting service's deployment form use exactly:

| Setting | Value |
| --- | --- |
| Repository | `joss16z/landmark-channel-toolkit` |
| Branch | `feature/live-floorplate` |
| Main file / entrypoint | `app.py` |
| Python | 3.11 or 3.12 |
| Python dependencies | root `requirements.txt` |
| Linux packages | root `packages.txt` for Report Tools OCR |
| Environment | `LANDMARK_PREVIEW_ONLY=1` |
| API Secrets | none; do not copy the API guide/key/sign into the host |

The local launcher is loopback-only. Community Cloud manages the server address
and port; follow the [separate Community Cloud guide](community-cloud-internal-test.md).
A generic approved internal platform may need IT to set a launch command like:

```sh
LANDMARK_PREVIEW_ONLY=1 python -m streamlit run app.py --server.address=0.0.0.0 --server.port=8501
```

Use that binding **only behind company-approved HTTPS, authentication and network
access restrictions**, with forwarded WebSocket support. Keep CORS/XSRF protection
and viewer restrictions enabled. Invite only approved staff. If the hosting option
only offers public access or has not been approved to process these uploads, stop
and use the local option instead. No paid service has been selected or introduced.

## Blockers and troubleshooting

- **Shared hosting:** viewer authentication/access policy and company permission
  to process artwork on Community Cloud must be verified before uploading there.
- **PDF rejected:** use the exact original six-page file; checksum/source evidence
  validation intentionally rejects altered or reordered PDFs.
- **No artwork yet:** upload both originals and confirm the review checkbox. The
  source documents are deliberately absent from GitHub, not missing app dependencies.
- **Live inventory:** the real getProject response schema/visibility/completeness is
  still unverified. Do not disable preview-only mode or enable the API in this test.

This prepares an internal test workflow; it is not approval for production or
publication of Internal/External Floorplates. Main remains unchanged.
