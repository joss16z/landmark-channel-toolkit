# Separate Community Cloud internal test app

This creates a **new test app**, not an update of production. The existing
production app and main branch must remain unchanged. Do not open its Manage app
page to change its source, Python version, Secrets, URL or privacy settings.

## 1. Create a new app

Sign in at **https://share.streamlit.io/** using your authorized GitHub account.
Choose **Create app** / deployment from GitHub. Do not redeploy or edit the
existing `landmark-channel-toolkit-ags5bvbwunrlszuoeyddum` production app.

Enter exactly:

| Field | Value |
| --- | --- |
| Repository | `joss16z/landmark-channel-toolkit` |
| Branch | `feature/live-floorplate` |
| Main file path / entrypoint | `app.py` |
| App URL | Choose a **new** unused name, e.g. `landmark-floorplate-internal-test`; do not reuse production's URL |
| Advanced settings → Python version | **3.12** (3.11 is also supported) |
| Advanced settings → Secrets | **Leave empty** — no Arcanite credentials needed |

Select **Deploy** for this new app only. The repository contains code and the
115-Lot coordinate map, not the original confidential PDF or PPT. Before uploading
anything, complete the privacy checks below. If GitHub access to the repository,
Python selection or deployment permission is unavailable, stop and ask your
workspace administrator; do not change production to work around it.

**Before starting:** read production's source settings, without editing them, and
confirm it does not track `feature/live-floorplate`. If it does, stop: feature-
branch updates would also affect that deployment, so the owner must resolve the
source separation before treating this as an isolated staging workflow.

## 2. Restrict the test app before uploading artwork

In the **new test app's** settings or Share dialog, choose the private option
(e.g. **Only specific people can view this app**) and invite only approved staff
using their individual email addresses. Do not select public access or a public
share link. If your workspace does not provide this option, or its private-app
quota is exhausted, **stop and do not upload company files**. Ask IT for an approved
private deployment option; do not change production's privacy or introduce a paid
service without owner approval.

Verify access using the new test URL:

1. A signed-out/incognito visitor must be blocked from the app by sign-in/access
   controls. They must not reach its file upload controls or generated downloads.
2. An approved invited user can sign in and open the app.
3. The original production URL remains a separate app with its original settings.

A public GitHub repository, an obscure test URL or a feature branch is **not**
access control. A private setting also does not replace company approval to process
artwork on Community Cloud. If company policy has not approved this host, test
locally using [the Windows guide](internal-preview-deployment.md) instead.

The current account's actual private-app permissions/quota and dashboard labels
could not be verified from this execution environment. Public documentation
access was blocked by the network proxy (403). Treat the access test above as a
mandatory deployment acceptance gate, not an assumed platform guarantee.

## 3. Test Live Floorplate

After private access and company hosting approval are confirmed:

1. Open the **test** URL; choose **Live Floorplate** in the sidebar.
2. Leave **Source PDF preview (synthetic inventory)** selected. The deployment
   defaults to preview-only; no API option or credentials are required.
3. Upload `Hurstville_Floorplate_E-LVL8 - Draft Strata Issue A 27.07.26.pdf`
   under **Original architectural PDF floorplates**. Upload just that original
   six-page file, including the cover, without re-saving or reordering it.
4. Upload `Hurstville Rise-Floor plate & Pricing-20261002.pptx` under
   **Reference Landmark PowerPoint**.
5. Leave **Optional replacement source-checked mapping JSON** empty.
6. Check the source/mapping review confirmation. Select each of the five groups:
   Ground; Level 1; Levels 2–4; Levels 5–7; Level 8.
7. Download **five-slide as-of PowerPoint**. Verify five slides, sage rounded
   boxes, shared Available units, and unchanged apartment artwork. All TEST
   units, statuses, bedroom/bathroom values, areas and prices are synthetic.
8. After expiry, use **Refresh synthetic source-layout preview**. Check both
   existing Report Tools remain accessible. Do not publish test PPTs as inventory.

Uploads stay in the current server session and may need repeating after a restart.
Do not upload originals using GitHub's Add file button or commit them to Git.
Do not enter the API key or signed URLs into the test app or its Secrets.

## Automatic dependency/configuration files

Community Cloud installs the root `requirements.txt`:

`streamlit>=1.37,<2`, `pandas`, `python-pptx`, `openpyxl`, `xlrd`, `lxml`,
`html5lib`, `beautifulsoup4`, `pdfplumber`, `PyMuPDF`, `pytesseract`, `Pillow`.

It uses the existing root `packages.txt` for **`tesseract-ocr`**, needed for existing
Report Tools OCR. Source PDF preview itself does not use OCR. No extra dependency
file, new entrypoint or API Secret is needed.

The committed `.streamlit/config.toml` enables headless startup, CORS and XSRF
protection, disables usage-statistics collection and retains the existing default
200 MB upload allowance so Report Tools' upload size is not reduced.
It does **not** force a loopback address or port: Community Cloud manages those.
Do not change these security settings or paste Python/configuration code into
Secrets. The default preview-only application policy blocks live synchronization
server-side even if a saved session previously selected the API mode.

## Verification and remaining blockers

Repository tests cover the preview default, API guard, all five group choices,
115-Lot map, five-slide export and existing Report Tools. A loopback Streamlit
startup/health test is run locally; that is not proof of a deployed Cloud instance.
The actual new Cloud app URL, authenticated viewer restrictions, dependency build
logs and company approval must be verified by the owner in the Cloud dashboard.
No hosted test app has been created from this execution environment, and production
has not been managed or redeployed. Real API response validation remains separate
and does not block the synthetic source preview.
