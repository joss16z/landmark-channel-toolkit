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
- Verified project ID, schema paths, response scope and authentication configuration
  still require the authorized operator. API signs stay in server-side Secrets.
- Owner visual approval remains required for publication. This reproduces the
  five-slide structure/style, not every historical box coordinate, since that PPT
  labels only its then-Available subset. New coordinates cover every footprint.
- PDF artwork is rasterized into PPT; information boxes remain editable. Browser
  font fallback is Arial when Aptos is unavailable. Native PowerPoint rendering
  and final company-device visual approval are still pending.
- No missing Lot becomes Sold. No official Internal/External Floorplate, Dropbox
  file or inventory is automatically modified or published.
