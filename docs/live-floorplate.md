# Live Floorplate operator and development guide

Run the existing app with `streamlit run app.py` after installing `requirements.txt`.
Run the complete suite with `python -m unittest discover -s tests -q`.
The existing Floorplate Update Report and PPT Price Update tools are unchanged.

## Three modes

1. **Synthetic demonstration**: generated fake drawings and fake apartments; no
   company files or API access needed. Demonstrates the five floor groups.
2. **Source PDF preview (synthetic inventory)**: upload the supplied original
   six-page Hurstville PDF and reference five-slide PPT. Check the review box.
   The bundled mapping matches this precise PDF by checksum. All 115 Lots are
   assigned conspicuous TEST units and fake values to stress-test layout only.
   Use manual refresh, select any of the five groups, and download a five-slide
   sample PPT. No Secrets or network connection is needed for this mode.
3. **Configured Agent Portal API**: same uploaded artwork and checked mapping,
   but only a successful validated server-side sync can provide real labels.
   This mode remains blocked until the actual response contract is verified.

Use company-approved hosting for real uploads. Uploaded documents and all snapshots
stay in the current server session, not Git. Source artwork is never overwritten.
The provided PPT is read for styling only, not for historical prices or current
availability. Backgrounds must contain no historical price labels. The originals
are not bundled in the repository; the checked coordinate mapping is bundled.

## Source mapping

The complete `live_floorplate/hurstville_mapping.json` is version 2, covering
Lots 1–115 and 57 source footprints. PDF page indices are zero-based: Ground=1,
Level 1=2, Levels 2–4=3, Levels 5–7=4, Level 8=5; page 0 is the cover.
Each record contains the source number bbox/page, floor, group, common box ID,
normalized label rectangle and normalized leader anchor. Stack members share the
same anchor and box coordinates. Lot is the matching key; never infer it from Unit.

`verify_source` checks the exact original PDF, every source number, stack order,
115-Lot coverage, 57 footprints, shared coordinates, and box/artwork collisions.
Edited or replaced assets must be reviewed and their mapping reverified; a checksum
mismatch is an error, not an invitation to change only the stored checksum.
The optional replacement mapping uploader accepts source-evidenced version 2 only.
Version 1 remains internal to the generated synthetic demonstration. The template
is intentionally unverified and cannot authorize live rendering.

Source verification does not approve publication; `owner_publication_approval`
is false. The operator's review checkbox confirms this run's source inspection.
See [validation evidence](live-floorplate-validation.md) for coverage and checks.

## Information boxes and export

Both online view and export use Ground, Level 1, Levels 2–4, Levels 5–7, Level 8.
Each footprint has one rounded light sage box, containing only Available apartments.
A group may contain multiple floors' units; unavailable members are omitted entirely.
Bedroom/bathroom and Internal + External fields accompany each current Unit/price.
Shared headings are compressed only when all member values agree. Different data
is retained rather than assumed identical. Text fitting is shared by HTML and PPT;
if it cannot remain legible, rendering fails rather than clipping a price.

Export has exactly five slides, PDF backgrounds filling reference-size slides,
editable boxes, source leaders and as-of footers. New labels are outside the
apartment artwork. The source PPT's historical positions covered only its then-
Available subset; the new mapping has positions for every footprint. This is not
a promise of pixel-identical historical presentation. Owner visual approval is
required before publishing. Source PDF rasterization preserves the original view
but does not create editable architectural vectors in the exported PPT.

## API configuration: still unverified

Use `.streamlit/secrets.toml.example` to configure only the approved server.
It is disabled and contains placeholders. Put `agencyid`, `sign` and the verified
Hurstville `projectid` in `[arcanite.query]`. Never put them in browser inputs,
Git, logs, screenshots, PRs or chat. No live login, endpoint discovery or live API
request was performed here. The official request guide has been reviewed and project ID 907440240 is confirmed.
No project-list response or getProject payload has yet been supplied. The redacted
response is being prepared by the owner. See [the exact request contract and
unverified response fields](official-arcanite-api.md).

The account is a standard Agent Portal account, not an administrator/backend account.
Use only the verified read-only request and inventory legitimately visible to it.
Never expand permissions or probe for hidden inventory. Signed query values stay
server-side, and redirects/cross-origin pagination are rejected before forwarding.

The adapter is configurable rather than asserting unobserved response fields:

- `fields` maps the actual row fields to lotnumber, propertystatus, unit, bedrooms,
  bathrooms, Internal, External and price. `project` is optional only when
  `project_context_verified=true` and a nonempty projectid binds the verified
  request to Hurstville. If response-level project identity exists, configure
  `scope_path` to validate it against the requested projectid.
- `rows_path` is the verified path to the list. No property-level project ID is
  assumed or required when using the verified single-project request context.
- `completeness=total_field` reads a **real**, verified `total_path` and checks
  unique scoped Lots and extracted rows against it, rejecting changes during paging.
- `completeness=expected_lot_set` requires a reviewed explicit `expected_lots` list.
  It needs no response total field. It accepts the result only if the endpoint
  returns exactly that complete set, including unavailable properties. For a
  full-project 115-Lot contract, configure all 115 Lots only after verification.
  Do not pretend this proves completeness of an intentionally restricted subset.
- `pagination=single` is allowed only for a verified complete single response;
  `next_link` in the generic adapter requires a real `next_path` terminating in explicit
  null, but the official getProject profile rejects it because the guide documents
  paging only for project lists. Offset,
  cursor, nested-array or POST-based contracts require a separately verified adapter.
  GET-only does not itself prove read-only semantics.

Neither completeness option is currently verified for Arcanite. Unknown fields,
missing fields, access errors, duplicates, inconsistent scope, malformed values,
partial results and unknown statuses all fail closed. Missing Lots are never Sold.
No schema, project count or pagination fields are invented for the live service.

## Freshness and publication

Refresh invalidates previous data before network access. A failed refresh cannot
retain old labels. Successful data carries a UTC sync time and expires after five
minutes. UI checks every five seconds, with an iframe expiry timer; PPT generation
also checks freshness. A downloaded PPT is an as-of artifact, not a perpetual feed.
There is no shared inventory cache, automatic publication, official PPT update,
Dropbox connection, paid API, merge or deployment in this feature branch.
