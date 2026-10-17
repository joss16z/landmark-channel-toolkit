# Hurstville Rise: complete source mapping and five-group rendering

Work remains on `feature/live-floorplate`, updating existing Draft PR #1 only.
No main merge, deployment, real API call or publication was performed.

## Mapping evidence and coverage

The PDF uses standalone, often space-separated numbers, not literal “Lot” text.
The earlier literal-word search was insufficient; it has been replaced by numeric
line extraction and source bounding-box verification. No apartment Unit number
is used to determine a Lot's floor or physical footprint.

| Source group | PDF page (1-based) | Lots | Coverage | Shared footprint anchors |
| --- | --- | --- | --- | --- |
| Ground | 2 | 1–4 | 4/4 | 4 |
| Level 1 | 3 | 5–18 | 14/14 | 14 |
| Levels 2–4 | 4 | 19–60 | 42/42 | 14 |
| Levels 5–7 | 5 | 61–105 | 45/45 | 15 |
| Level 8 | 6 | 106–115 | 10/10 | 10 |
| Total | cover excluded | 1–115 | **115/115** | **57** |

`live_floorplate/hurstville_mapping.json` records every Lot, floor, group, common
box ID, label rectangle, leader anchor, original PDF number bounding box and source
PDF SHA-256. Floors 2–4 and 5–7 follow the printed number stacks top-to-bottom in
the ascending floor order shown by the source titles. Per-floor coverage is
4, 14, 14, 14, 14, 15, 15, 15 and 10. This comes from PDF callouts, not Unit IDs.

The mapping is marked source-verified after all five source pages were visually
inspected, all 115 number bounding boxes compared with the source, all 57 stack
anchors checked, and label coordinates tested against the apartment-artwork
bounds and against other boxes. Runtime verification repeats checksum, coverage,
stack ordering, shared geometry, source evidence and overlap checks. A changed
PDF or altered evidence fails closed. `owner_publication_approval` remains false:
technical source verification is not the owner's approval to publish prices.
The coordinates/evidence are committed as requested; original documents and
historical price data are not committed.

## Five-slide and online validation

Export now creates **exactly five slides**, in the original floor-group order.
The original PDF page fills each slide, preserving its drawing, branding and
aspect ratio at the reference PowerPoint dimensions. New labels are editable
rounded rectangles with the reference light sage fill, font and color. Labels
sit around the apartment artwork, with leaders to the source callouts. Grouped
Available apartments share one box. Shared bedroom/bathroom/area headings are
printed once; differing configurations retain each apartment's own information.
Only the current supplied snapshot supplies prices; reference PPT prices are
never copied. Text is wrapped/scaled conservatively and export fails if it cannot
fit legibly; HTML uses the same fitted content. Missing bathroom data also fails.

A separate private integration check used the supplied originals and all 115
**synthetic** Available records. It produced five slides and 57 boxes, with no
box-to-box collision or overlap with the apartment-artwork bounds. Original files
were unchanged. Actual Chromium checked all five generated HTML groups at a
1500-pixel viewport: box counts **4, 14, 14, 15, 10**, with **zero overflowing boxes**.
The source PDF pages and generated browser previews were visually inspected.

Streamlit now selects the same five groups. Source-PDF preview uploads the original
PDF and reference PPT, uses the bundled source-checked mapping, and renders all
115 synthetic records without Secrets or API access. The user-facing warning and
PPT footer identify the data as synthetic; they are not real inventory. Both
existing Report Tools remain functional. Refresh/five-minute expiry safety remains.

## Tests

Command: `.venv/bin/python -m unittest discover -s tests -q`.

**29 tests passed; zero failures, zero errors, zero skips.** The suite includes
source-evidence tamper rejection using synthetic PDFs; complete mapping/floor
coverage; shared anchors; source/box collision rejection; five-slide generation;
multi-unit Available-only boxes; bedroom/bathroom/area/price content; missing fields;
API paging/scope/completeness failures; stale/failed refresh; Streamlit five-group
navigation; and the existing report-generation and Report Tools navigation checks.
Synthetic PDF tests do not embed the company documents. Source-specific checks
above were run separately against the private supplied originals.

## Remaining blockers and limitations

- **Actual getProject schema validation remains pending the redacted response.**
  The API adapter has not been validated against or called on the live service.
  The supplied signed URLs establish an endpoint, not a response schema.
- The adapter no longer requires property-level project identity when a verified
  single-project request context is configured. A total field is used only when
  it actually exists. An alternative exact, reviewed Lot-set completeness proof
  supports responses without totals. For all 115 project Lots, that mode requires
  all 115 including unavailable properties; it must not be used if the authorized
  endpoint intentionally returns a subset. Unsupported/unknown completeness stays
  blocked. No pagination or metadata fields are invented for the real service.
- Project ID 907440240 is now owner-confirmed. Schema paths, response scope and
  authentication configuration still require verification. API keys stay in Secrets.
- Owner visual approval remains required for publication. This reproduces the
  five-slide structure/style, not every historical box coordinate, since that PPT
  labels only its then-Available subset. New coordinates cover every footprint.
- PDF artwork is rasterized into PPT; information boxes remain editable. Browser
  font fallback is Arial when Aptos is unavailable. Native PowerPoint rendering
  and final company-device visual approval are still pending.
- No missing Lot becomes Sold. No official Internal/External Floorplate, Dropbox
  file or inventory is automatically modified or published.

## Official request documentation follow-up

The official Landmark integration guide has now been reviewed. GET methods,
endpoint paths and required agencyid/projectid/sign parameters are documented;
sign is the fixed API key. Hurstville project ID **907440240** is owner-confirmed.
The official request builder enforces these parameters and restricts documented
project-list filters/paging to getProjectList. Live sync remains disabled because
property schema, status/price visibility and completeness are still unverified.

See [the request/schema review](official-arcanite-api.md) for the exact outstanding
response fields and vendor guarantees. An offline schema inspector supplies
structural diagnostics without raw property or credential values. The mapping
and five-group rendering continue unchanged; the API guide PDF is not committed.

Follow-up suite: `.venv/bin/python -m unittest discover -s tests -q` — **40 passed,
zero failures/errors/skips**. Added checks cover documented parameter dependencies,
missing keys/IDs, forbidden endpoints/guessed property paging, redaction of signed
transport errors, configured error envelopes, and offline schema/coverage diagnostics.
All previous 115-Lot mapping, five-slide rendering and Report Tools tests still pass.

## Internal Community Cloud test-deployment preparation

The feature branch now defaults to **preview-only**, hides the live API option,
clears a previous API-mode selection and blocks synchronization server-side. No
Secrets or credentials are required. Source PDF preview is the default inside
Live Floorplate, with both existing Report Tools still accessible.

The Cloud entrypoint remains root `app.py`; dependencies remain root
`requirements.txt` and Linux `packages.txt`. `.streamlit/config.toml` retains
CORS/XSRF protections, disables usage statistics and leaves bind address/port to
Community Cloud. The Windows local launcher explicitly binds loopback instead.
See the [separate Community Cloud checklist](community-cloud-internal-test.md).
No production app settings or main files were changed; no hosted app was created.
The owner must verify production tracks a separate branch, create a **new** test
app and verify private, invited-only access before uploading originals. Current
Cloud-account permissions/private-app quota remain unverified. Public Streamlit
docs access was blocked by the execution network proxy (403); no bypass was used.

Verification: **42 tests passed; zero failures/errors/skips**. A local Streamlit
startup returned HTTP 200 from its health endpoint. Headless Chromium then used
the actual app upload controls for the supplied original PDF and reference PPT,
confirmed live API controls absent, and downloaded an exactly five-slide sample
PowerPoint. That private loopback check is not proof of a hosted Cloud deployment
or Cloud viewer privacy. The test server was stopped after verification.

## Connection diagnostic regression validation

Added an explicit `Test Arcanite Connection` action independent of inventory
synchronization. Synthetic tests cover direct/nested server Secrets, conflicting
or missing configuration, 401/403 responses, denied redirects, foreign response
origins, non-JSON/malformed/oversized bodies, redacted failures, and button-driven
success/failure without inventory snapshot creation. Summary inspection includes
field names/types, Lot counts and fixed status counts only. A structural result
cannot certify live authentication, authorized inventory scope or completeness.

Validation: `.venv/bin/python -m unittest discover -s tests -q` — **50 tests passed**.
Existing PPT tests were aligned with the branch's previously updated area format
and wrapped Unit labels; five-slide and availability checks remain active.
No live endpoint was called, no credentials were used, and live synchronization
remains disabled in the default internal preview deployment.

### HTTP 403 audit follow-up

Compared `connection_test.py`, `official_api.py` and `api.py` with the official
request contract on guide pages 4–5. No endpoint/method/query-authentication
construction mismatch was found. Added network-boundary regressions for the
exact GET request and single encoding of synthetic reserved characters, plus a
403-before-schema test. The vendor/gateway rejection source remains unverified.

Full suite: `.venv/bin/python -m unittest discover -s tests -q` — **52 passed**.
No live request, credential inspection, production change or price-sync enablement
was performed. Approved-client comparison and vendor/IT log correlation are still
required to separate credential/project authorization failure from network policy.

### Owner-provided two-endpoint diagnostic

Added an independent, explicit `Test Project List Access` button using only the
official GET getProjectList with agencyid/sign. No optional filters, project
link following, pagination or response-envelope assumptions were added. Synthetic
regressions cover its exact transport parameters, private-value redaction,
unsupported endpoint rejection, HTTP 403 and independent UI invocation without
snapshot creation or price-sync enablement. HTTP 200 is not an authentication or
visibility guarantee. Full suite: **55 passed** using
`.venv/bin/python -m unittest discover -s tests -q`. No live calls were performed.
