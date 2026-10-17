# Official Arcanite API contract review — Landmark

Source: owner-supplied **Arcanite Third Party Integration – Landmark**, six pages.
The PDF contains credential material and stays outside Git. Its published request
specification is authoritative for parameters; it is not evidence of unspecified
response fields. The Response example headings contain no response schema.
No live API request, browser Network inspection, login, merge or deployment was used.

## Verified request facts

| Operation | Method | Required query parameters | Document pages |
| --- | --- | --- | --- |
| `/arc-service-api/v1/project/getProjectList` | GET | `agencyid`, `sign` | 1–4 |
| `/arc-service-api/v1/project/getProject` | GET | `agencyid`, `projectid`, `sign` | 4–5 |

Both operations use `https://service.arcnet.com.au`. `sign` is the fixed API key
specified by the guide; no hashing, HMAC or invented signature algorithm is used.
Hurstville Rise's request project ID is **907440240**, confirmed by the owner.
The template has this ID and GET configured, while `enabled=false` and
`contract_verified=false` continue to prohibit calls with an unverified response
contract. Actual agency ID/API-key values belong only in server-side Secrets.

`getProjectList` optionally accepts `searchkey`, `category`, `minprice`, `maxprice`,
`israngesearch`, `bedroomqty`, `bathroomqty`, `studymediaqty`, `carspaceqty`,
`pagenum`, `pagesize` and `sortby`. The request builder enforces documented pairs
and dependencies: page number/page size together; maximum price with minimum;
quantity filters with range mode; categories 01–04; range mode 0/1; sort 1/2.
Blank optional parameters are omitted. These are **project-list filters**, not
property response field definitions. They do not establish property pagination.

The official getProject profile rejects additional/guessed query parameters and
undocumented property next-link pagination. The generic synthetic-test adapter
still tests pagination safely, but that is not evidence the real endpoint supports it.
The request builder allows only these two read endpoints; guide contact/enquiry
write operations are not implemented.

## Visibility and completeness: not yet established

The guide says projects must be set to **Show on Website** for getProjectList.
This means API project visibility may differ from the user's Agent Portal views.
It does not grant administrator access or authorize the connector to change that
setting. It also does not say whether getProject includes every project Lot,
only published properties, only Available properties, Held/Pending/Sold entries,
or prices and areas for any of those entries.

Therefore it is currently **unverified whether getProject returns a complete
property list**. Do not infer Sold, zero availability, or whole-project coverage
from a missing record. The 115-Lot PDF map proves architectural coverage, not API
visibility. No unavailable property or historical reference-PPT price is displayed.

## Exact additional response information needed

Supply a redacted successful getProject JSON for **907440240**, preserving keys,
object/list structure, types and nulls. Replace confidential property values with
sample values and remove API keys, signed URLs, tokens and unrelated project data.
For completeness validation, retain the Lot set locally; use the structural
review tool below to share counts/coverage rather than real inventory.

| Required confirmation | What needs verification |
| --- | --- |
| Response envelope | Success/error indicators, their values/types, and HTTP failure behavior; a redacted error response or vendor error contract |
| Property collection | Exact object path to properties; whether properties can be omitted, nested, filtered or paginated |
| Lot identity | Actual path/type of `lotnumber`, missing/duplicate handling and leading-zero conventions |
| Unit identity | Whether `unitnumber` exists and its type/display format; currently an owner-specified candidate |
| Availability | Actual `propertystatus` field/type, exact enumerated values, and which statuses this agency key may receive |
| Current price | Actual price field/path/type, currency and whether it is the current agent-visible sale price rather than another price or an undisclosed/null value; `price` remains a candidate |
| Areas | Exact Internal/External field paths, units, types and null behavior; no response area names are documented |
| Bedroom/bathroom | Exact response paths/types; list-filter names do not prove response field names; any study/media configuration needed for the label |
| Project scope | Whether root metadata identifies project 907440240; no per-property project ID is required or assumed |
| Completeness | Vendor confirmation whether a single response includes all properties authorized for this project, including status restrictions; actual total/page/end fields only if they exist |
| Freshness | Whether response/update timestamps, cache age or snapshot versions exist and what they guarantee |

If the endpoint supplies all 115 Lots including unavailable properties, an exact
reviewed Lot-set proof can validate completeness without inventing a count field.
If it intentionally returns only a subset, that mode must fail. We instead need
a documented authorized-view total or another vendor-supported completeness proof.
If pagination or a separate property/listing endpoint is needed, the **official
endpoint/method/parameter and response documentation** is needed; no alternative
endpoint will be guessed. A getProjectList project count cannot stand in for a
getProject property count. No further endpoint is currently established as necessary.

## Offline structural review in the approved environment

Save an authorized response using your approved API client, **outside the Git
checkout**. Do not put a signed URL or key into shell arguments, history or logs.
Then run (substitute your protected local response path):

```powershell
python -m live_floorplate.schema_review --response "C:\ApprovedPrivate\getProject.json"
```

This tool makes no HTTP requests. It prints candidate property-array paths, field
names/types, array counts, duplicate/invalid Lot counts, fixed-category status
counts, and whether exactly Lots 1–115 are present. Metadata candidates show paths
and types only. It does not print unit values, prices, raw statuses, credentials,
URLs or raw JSON. It always reports `endpoint_completeness_verified=false` because
structure inspection cannot establish server visibility/contract guarantees.
A redacted response plus the vendor's scope/completeness statement is still needed
before mapping actual fields and enabling live sync.

## Implemented safe behavior

All HTTP is server-side; secrets enter only the documented query parameters.
Undocumented endpoints/parameters, redirects and cross-origin pagination are
refused. Transport errors are replaced with fixed exceptions, suppressing signed
URLs in normal tracebacks. Only HTTP 200 JSON within the response-size limit can
reach validation. Verified success markers can be configured without assuming a
particular error envelope. Wrong/missing fields, malformed prices/areas, missing
bathrooms, duplicate Lots, inconsistent scope, unsupported completeness and partial
results fail closed. A failed refresh removes prior labels before requesting data.

The existing 115-Lot mapping, five-slide export and five-group source-PDF preview
remain independently usable with clearly synthetic data while response verification
is pending. No source document or confidential response is committed.

## Manual connection diagnostic (no inventory synchronization)

On the Live Floorplate page, expand **Arcanite connection diagnostics** and click
**Test Arcanite Connection**. No request occurs merely by opening the page. Use
this only in a company-approved deployment with restricted viewer access: an
invited viewer who can access this page can trigger the server-side request.
Do not add credentials to a public/unrestricted test deployment.

In Streamlit's server-side Secrets editor (or an ignored local
`.streamlit/secrets.toml` on the approved computer), configure:

```toml
[arcanite]
agencyid = "REPLACE_LOCALLY"
sign = "REPLACE_LOCALLY"
projectid = "907440240"
```

These are placeholders, not credentials. The existing `[arcanite.query]` form
also works; conflicting direct/nested values fail before any request. Do not
paste a signed URL, upload Secrets, or commit the local Secrets file. The test
uses only these three parameters, the fixed official getProject endpoint, GET,
a 20-second timeout and the existing 10 MB response limit. It sends no additional
headers, follows no redirects, caches no response and saves no diagnostic files.

The display contains only HTTP status (or unavailable), field names/types, Lot
counts and fixed status-category counts. Unknown status values are counted as
`unknown`, never displayed. Raw JSON, Lot identifiers, unit numbers, prices and
credential values are not displayed or logged. Errors are fixed messages, not
transport exception text. HTTP 200 and a property-shaped array are only
structural observations: they do not prove authentication, all-status visibility,
completeness or price validity. An absent/changed property array fails the test.

This diagnostic neither calls inventory `sync()` nor creates a snapshot. Keep
`LANDMARK_PREVIEW_ONLY` at its safe default; do not change `enabled` or
`contract_verified` to run the diagnostic. Failed tests leave live price sync
disabled. Actual response schema, authorized visibility and completeness remain
pending owner review; this change was tested using synthetic responses only.

## HTTP 403 request audit and approved diagnosis

Audit against the owner-supplied official guide, pages 4–5:

- Endpoint: `https://service.arcnet.com.au/arc-service-api/v1/project/getProject`.
- Method: GET; no request body. Required query keys are exactly `agencyid`,
  `projectid`, `sign`. getProjectList filters/paging are not sent.
- `sign` is the guide's fixed API key, not a calculated signature or a Bearer
  token. Browser Agent Portal login cookies are not this API's documented
  authentication mechanism. Do not introduce a guessed signing algorithm,
  browser Cookie, Referer or administrative credentials to address a 403.
- `connection_test.py` resolves direct/nested server Secrets;
  `official_api.py` checks the documented parameter contract;
  `api.py` sends the explicit GET with redirects disabled. `urlencode` encodes
  query values once, and URL fragments cannot consume parameter values.
  Supply the **raw** key in Secrets, not an already percent-encoded key, a full
  signed URL, an `Authorization` value, or a `sign=` prefix. The builder trims
  surrounding whitespace, but it cannot detect an expired or wrong key.
- Synthetic tests exercise the complete path down to the urllib Request,
  including reserved characters (`+`, `/`, `&`, `=`, `%`, `?`, `#`) and Unicode,
  both Secrets forms, exact keys/method/body, timeout and redirect handler.
  No request-construction defect was found in this audit; no speculative
  transport/authentication change was made.

An HTTP 403 is retained from the HTTP transport before schema inspection. A
missing/conflicting Secrets configuration fails before network access with no
HTTP status; malformed JSON/unknown schema cannot manufacture an HTTP 403.
However, HTTP status alone **does not identify the responding layer** or prove
that authentication failed. The vendor API, a reverse proxy, WAF, egress proxy
or another gateway can return 403. Wrong/expired credentials, agency/project
permissions and source-IP policy are possibilities, not confirmed diagnoses.
The urllib client normally has a Python User-Agent rather than a browser one;
a WAF may apply client-header rules, but the guide establishes no extra header
requirement. Do not spoof a browser or bypass access controls to test a theory.

### Next steps for the owner and company IT

1. In the **test app's private server Secrets settings**, locally check the
   agency/key pairing, current key and project ID `907440240`. Do not post the
   values or signed URL. A successful Agent Portal browser login does not prove
   that this separate third-party API key has project access.
2. With explicit company approval, compare the same documented GET and key on
   an approved work computer using the same Python connector and the company's
   approved API client. Disable client request/response logging/history exports;
   keep results and credentials local. Do not send raw responses back to GitHub
   or Codex. This audit did not perform any live comparison.
3. Interpret the comparison cautiously:

   | Observation | What it narrows down; what remains unproven |
   | --- | --- |
   | Both approved clients fail locally | Ask Arcanite to check key activation, agency/project authorization and API policy. A common network restriction is still possible. |
   | Approved local connector succeeds, Cloud test fails | Investigate Cloud egress IP/region, proxy and WAF policy with IT/vendor. It is not evidence of a signature-algorithm bug. |
   | Approved API client succeeds but Python fails on the same computer | Compare request method, query encoding and approved headers/proxy settings locally. Have IT/vendor identify the differing rule before any code change. |

4. Give Arcanite/IT only the UTC attempt time, non-sensitive endpoint path,
   HTTP status and environment (approved PC or internal Cloud app); they can
   correlate their server/WAF logs and provide a sanitized rejection reason.
   Ask whether the key is active for this agency/project, whether the API is
   enabled, whether source-IP allowlisting is required and which layer rejected
   the call. Never share the key, query URL, cookies or inventory in a ticket.
   Do not disable a proxy/WAF or change networks to evade a restriction.

The app discards error bodies and does not log request URLs or exception text.
Because this official protocol carries the key in the query, IT/vendor must
also ensure infrastructure/access logs redact query credentials; application
redaction cannot guarantee third-party logging behavior. Do not enable live
price synchronization based on a successful connection diagnostic alone.
