# Live Floorplate — first working version

The existing Streamlit app and both active Report Tools remain in place. Run:

```sh
python -m venv .venv
# Activate the environment using your platform's normal command.
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
streamlit run app.py
```

Choose **Live Floorplate** in the sidebar. The default demonstration is generated
from synthetic domain records and synthetic rectangular drawings. It is not an
Arcanite response, real inventory, actual Hurstville placement or approved Landmark
artwork. It demonstrates Ground through Level 8, reusing PDF pages for Levels 2–4
and 5–7, and generates nine PowerPoint slides from the same snapshot.

## Existing functions reviewed and reused

`app.py` provides unit/price formatting, PDF ingest and report/update UI functions.
They remain unchanged; importing app helpers into a renderer would execute the UI.
`floorplate_update_report_v8.py` provides `normalize_unit` and colored Unit-label
recognition, reused for accessible labels and reference PowerPoint styling. Its
`is_available` accepts aliases such as “yes” and “a”, so it is deliberately not used
for the strict API rule. The new page accepts only `propertystatus == "Available"`;
Held, Pending and Sold produce no new information boxes or prices. Unknown status
or invalid required Available data fails the complete snapshot.

## Pending source inputs — no invented live contract

The owner supplied the original Hurstville Rise PDF and reference PowerPoint;
they were inspected outside Git and are not committed. The PDF has a cover plus
five layout pages; the PowerPoint has five layouts with grouped shapes. No
explicit Lot references were found in their editable text. Confirmed Lot positions
and a redacted getProject schema remain pending. Exact visual parity and live
API interoperability are **unverified**.

The configured page uploads these assets into the current Streamlit session only.
Use an approved runtime: file uploads and secrets are available to that server,
so do not upload company documents to an unapproved public/Cloud deployment.
The app does not write uploads to the repository. Keep local files under ignored
`private/` or outside the checkout. Original PDF/PPTX bytes are never overwritten.
Only their rendered architectural background and extracted label style are used.
PDF backgrounds must have no embedded old prices/availability labels; the operator
must confirm this before use. The app cannot automatically verify that assertion.

## Explicit layout mapping

Start from `live-floorplate-layout.template.json`. `pages` maps level strings 0–8
to **zero-based** page indices in the uploaded PDF sequence. The template uses
the supplied six-page PDF: Ground=1, Level 1=2, Levels 2–4=3, Levels 5–7=4,
Level 8=5; page 0 is the cover. These are background bindings, not Lot mappings. Multiple PDFs are
combined in memory in upload order. Reverify indices if the upload sequence changes;
levels 2/3/4 must share one page and 5/6/7 another. No Lot-to-floor relationship is
inferred. `positions` contains entries shaped as:

```text
{"lotnumber": <confirmed lot>, "level": <integer 0..8>,
 "x": <fraction>, "y": <fraction>, "width": <fraction>, "height": <fraction>}
```

Coordinates describe the information box relative to the original PDF page,
with origin at its top-left and bounds 0..1. Every Available lot needs exactly one
confirmed position; unknown or duplicate mappings block both views and exports.
Use owner-approved label positions and sizes that fit the text without covering
important drawing details. Set `verified` true only after review. `lotnumber` is
the match key; unit numbers, record order, and numerical Lot sequences do not
choose the floor or layout. Leading-zero numeric lots share a canonical key;
duplicates within the project fail validation. Unknown extra positions are inert.

## API configuration and Agent Portal permissions

Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` only on the
approved server, or enter equivalent settings in approved Streamlit Secrets.
The example is disabled and contains the supplied public endpoint path, but no
authentication values or invented live response shape. Configure `agencyid`,
`sign` and the verified Hurstville `projectid` in `[arcanite.query]` locally;
the signed URL supplied in chat is never committed or requested during development.
`getProjectList` discovery is not implemented or called: the project ID still
needs authorized verification. Do not paste credentials, cookies, tokens or price lists into
GitHub, logs, browser code, chat or PR descriptions. No live authentication was
performed during development. No authentication/MFA bypass or discovery exists.

An authorized operator must establish:

- The exact getProject URL, read-only method and legitimately available account
  scope. This is a standard Agent Portal account, without admin/backend privileges.
- Actual field names and paths for project, `lotnumber`, `propertystatus`, unit,
  bedroom configuration, Internal Area, External Area and current displayed price.
- The all-status row count for the selected authorized project/view and its
  pagination contract. A count of Available-only rows is not interchangeable.
- The approved server-side authentication header, renewal procedure and runtime
  authorization; no credentials are requested through the Streamlit frontend.

The first adapter supports **verified GET requests** only, dot-separated object
paths, a list of row objects, an integer complete total, and either a single page
or a same-origin `next` link ending in explicit null. Empty `rows_path` means the
payload itself is the list, but an independent verified total is still required.
Other schemas, cursor/offset pagination, project metadata outside each row,
POST-only read APIs, cookie/browser-bound authentication and nested row arrays
need a separately reviewed adapter after the real contract is supplied. Do not
change the method or guess a backend route merely to make a request succeed.
GET alone cannot prove read-only semantics; the owner must verify the endpoint.

Set `enabled`/`contract_verified` true only after verification. Endpoint and headers
stay server-side. Redirects, cross-origin/repeated pagination, non-JSON responses,
401/403 failures, missing/moving totals, incomplete rows, duplicate Project+Lot keys,
unknown statuses and mixed projects cannot produce a usable snapshot. There is
no automatic network call on page load; synchronization requires the button.
No discounted or adjusted price is calculated. Areas remain Internal + External.

Completeness is restricted to the Agent Portal's authorized project/filter view;
it never proves unrestricted company-wide inventory. Missing properties are never
classified Sold. There is no publication or comparison against official inventory.

## Freshness, UI and export

A failed refresh removes the previous snapshot **before** networking. Successful
snapshots show a UTC last-synced timestamp and expire after five minutes. The view
checks every five seconds, with a browser timer hiding labels at expiry. No cache
is shared between users. PowerPoint generation also checks freshness and includes
an as-of timestamp on every slide; a downloaded file is a historical artifact,
not a perpetually current price feed. Download controls disappear on expiry.

HTML and PowerPoint consume the same immutable snapshot, Lot mapping and style.
The PDF drawing remains a raster background; information boxes are editable PPT
shapes. Reference solid RGB fill and first text-run font/size/color and reference slide dimensions are reused.
Grouped labels are inspected recursively.
Theme colors, complex multi-run style fidelity and grouping/connector-line copying
are not implemented; unresolvable reference styling fails safely. The synthetic
sample uses yellow/black Arial only as a demonstration. Owner visual review of
actual files is required before use or publication. Export never changes official
PPTs, writes to Dropbox, or publishes anything. No paid API/services are introduced.

## Review and testing

Tests use synthetic domain records and explicitly synthetic adapter contracts;
no fixture asserts a real Arcanite response format. They cover exact availability,
Lot matching, completeness, signed-query separation, prices, unsafe pagination, expiry, shared pages,
PowerPoint contents/style, and the existing v8 report-generation path. Streamlit
AppTest checks both existing Report Tools and the new sample page. Native browser
visual parity, real Hurstville inputs and live API integration remain pending.
