# First-version validation and manager handover

Branch: `feature/live-floorplate`, based on main `e990cb1`.

Implemented an additive Streamlit page, immutable validated data model, explicit
lot-position mapping, read-only signed-query API configuration, responsive PDF
background view, reference label styling and nine-slide PowerPoint generation.
Existing app changes are limited to sidebar navigation and the new page handler;
v8/v6 report code is unchanged. The existing devcontainer launcher now retains
Streamlit's default CORS/XSRF protections instead of disabling them.

`python -m unittest discover -s tests -q` in the installed project virtual
environment: **20 tests passed, no skips**. Coverage includes Available-only exact
status matching; omitted unavailable prices; duplicate, incomplete and mixed
project rejection; Lot matching independent of Unit/order; price formatting;
expired and failed refresh behavior; unsafe/repeated pagination; server-only
synthetic signed-query transport; configured mode disabled without assets;
shared layouts; PowerPoint contents and label style; existing report generation;
and Streamlit navigation through both active Report Tools and the sample page.
Streamlit AppTest uses synthetic data and makes no live API requests.

The owner-supplied PDF and PowerPoint were inspected outside the checkout.
Their five floor groups correspond; the PDF includes an additional cover page.
Editable source text does not provide explicit Lot references. Colored reference
labels use a consistent fill, and grouped shapes are traversed for style extraction.
A separate in-memory smoke check used those original PDF backgrounds and reference
styling with clearly synthetic labels: nine slides generated, reference slide
size preserved. No original documents, source prices, real inventories or derived
Lot mappings are committed. Original files were not changed.

## Still required before live or publication use

1. Owner-verified Hurstville Rise project ID and redacted successful getProject
   JSON, including complete-count/pagination semantics and all required field paths.
   Supplied endpoint URLs alone do not establish a response schema; neither endpoint
   was called. The project-list adapter/discovery is not implemented.
2. Owner-approved Lot-to-floor/label positions. Unit-only text in the source PPT
   does not justify guessing Lot numbers. Mapping completeness gates rendering.
3. Authentication values entered only in Streamlit Secrets in the approved runtime;
   owner/IT confirmation that this Agent Portal account may read this endpoint.
4. Visual review of the actual PDF backgrounds, box positions and text fit. The
   first reference label's RGB fill/font/size/color and slide dimensions are reused;
   theme styles and complex per-run formatting remain unsupported. PDF images are
   rasterized without altering their originals; new PPT boxes remain editable.
5. Confirm original backgrounds contain no old prices or availability labels.
   Text inspection alone cannot guarantee image backgrounds are free of labels.

The online view uses individual floors and shared background pages; export produces
one slide for each floor (nine), rather than reproducing the source's five multi-floor
slides and multi-unit text groupings. Snapshots expire after five minutes and
always display an as-of timestamp; a downloaded PPT is a historical snapshot.
A stable row total is not proof of an atomic price snapshot during concurrent
upstream edits; response version/timestamp guarantees remain unverified.
There is no automatic publication, Sold inference, Dropbox integration, merge or
deployment. API method/schema variants beyond the documented GET/next-link contract
must be adapted only after source verification. No paid services were introduced.
