# Learnings — People Directory WebPart

> General, reusable knowledge (SPFx / Graph / SharePoint / standards).
> Codebase-specific findings + issues live in `discoveries.md`. Newest at top.

## Build / tooling
- **SPFx 1.18 gulp needs Node 16/18** — Node 22 is rejected outright. Use `nvm use 18` (available) before any `gulp build/bundle/serve`. `tsc` alone runs on 22 but skips lint + scss codegen.
- The generated `*.module.scss.ts` lists the CSS-module class names TS can reference; reuse existing classes or keep new controls self-contained (Fluent + inline) to avoid regenerating it.
- Fluent 8 v8 has **no Pagination** in `@fluentui/react` core — build a small pager from `DefaultButton`/`IconButton`.
- This repo's ESLint enforces **`react/jsx-no-bind`** — no inline arrows in JSX; use `useCallback`/sub-components (see `LetterIndex`, `Pagination`).
- **A11y is a build gate here** — active nav item needs `aria-current`; keep boundary-disabled buttons focusable with `allowDisabledFocus`.

## Responsive / CSS (SPFx + Fluent 8)
- **Enforce touch targets on Fluent controls** without editing each TSX: in the SCSS module, nest `:global(.ms-Button){min-height:44px;min-width:44px}` (also `.ms-Dropdown .ms-Dropdown-title`, `.ms-SearchBox`) inside the web-part root class + `@media (max-width:640px)`. Compiles to `.peopleDirectory_x .ms-Button` — scoped to the web part, and its 2-class specificity beats Fluent's own base + any `styles={{root:{...}}}` merge-styles class. `min-height` overrides Fluent's `height:32px` (min-height wins when larger).
- **Caveat**: Fluent `Panel`/`Callout`/`Dropdown` callout portal to `document.body`, *outside* the web-part root — scoped `:global` rules don't reach them. Style those via the component's `styles` prop.
- **Prevent text overflow in a Fluent `Stack` row**: the text needs `min-width:0` (flex children default to `min-width:auto` and won't shrink) plus `word-break:break-word; overflow-wrap:anywhere`.
- SPFx media queries are **viewport-based**, not container-based — a web part in a narrow column on a wide monitor keeps desktop rules.
- Verify CSS responsiveness without SharePoint/Graph: inline the compiled `lib/**/*.module.css` into a static HTML harness (reuse the hashed class names from `*.module.scss.ts`), serve over `http://localhost` (Playwright blocks `file://`), and assert `scrollWidth<=clientWidth` + tap-box sizes at 320/360/375px.

## Workflow / process
- Give review/analysis subagents **read-only** tools. A review workflow once mutated a source file (removed a feature block) — always run `git diff HEAD` and rebuild before trusting a post-review tree.

## SharePoint / PnPjs — bulk writes & sync
- **PnP `$batch`** (`import '@pnp/sp/batching'` → `const [bsp, execute] = sp.batched()`) collapses many item add/update calls into one HTTP request. Batch **50–100 ops** (2 MB payload limit; batches are **not transactional** — inspect per-op results and retry failures).
- **Bulk upsert without a per-user existence query**: preload `Map<upn, itemId>` once via **keyset paging on the indexed `Id`** (`filter('Id gt <lastId>').orderBy('Id',true).top(5000)` loop) — scales past the 5,000 threshold (unlike `.skip()` offset paging). Route add-vs-update from the map; this also makes re-sync idempotent (no duplicate rows).
- **Update-by-ID does NOT hit the 5,000 view threshold** (only querying large unindexed sets does). So writes scale; the *read* to build the map is the part that needs keyset paging.
- **Throttling**: honor `Retry-After` (read `error.response.headers.get('Retry-After')`) on 429/503, else exponential backoff+jitter; **retry**, never silently drop. Process batches sequentially; don't spike concurrency.
- **Accurate item count >5000**: read the list's `ItemCount` property (`getByTitle(t).select('ItemCount')()`), not `.top(5000).length`.
- SharePoint has no unique-value guarantee unless `EnforceUniqueValues` is set; app-side dedupe (preload map + de-dupe input by key) is how this project prevents duplicate rows.

## SharePoint / PnPjs
- **5,000 list-view threshold is fixed** (per-query, not storage) and cannot be raised. Index every column you filter/sort on — at creation, while the list is small. Max 20 indexed columns. To read a whole large list: keyset-page on the indexed key (`Id gt <lastId>` ordered by `Id`) then filter/sort in memory. A bare `.top(5000)` is silent truncation, not a safe cap.
- **`PD_` field prefix** avoids SharePoint reserved-name collisions on custom columns.
- Code should bind to a column's **internal name**, never its display name.
- Prefer counting via list metadata (`ItemCount`) over fetching rows and taking `.length`.

## Microsoft Graph
- **`MSGraphClientV3` auto-prepends the `/v1.0` (or `/beta`) version prefix.** When reusing an `@odata.nextLink`, strip the version segment first or the path doubles up (see `GraphService.parseNextLink`).
- **`ConsistencyLevel: eventual`** header is required for `$count` and advanced query params (incl. `$search`).
- **The Graph JS client (`@microsoft/microsoft-graph-client`) does NOT URL-encode the query string** — `GraphRequest.createQueryString` concatenates `key=value` verbatim. So when building a raw endpoint string with special chars (e.g. `$search="displayName:x"`), **`encodeURIComponent` the value yourself**. No double-encoding risk since the client doesn't re-encode.
- **`$search` on `/users` is NOT true fuzzy** — tokenized/prefix, order-independent on displayName; other fields fall back to `startswith`. `$skip` unsupported → page via `@odata.nextLink` (max 999/page). Requires `ConsistencyLevel: eventual`.
- **Graph v1.0 bug**: an encoded ampersand (`%26`) in a directory-object `$search` returns `400`. Sanitize `&` out of the term (replace with space to keep token matching) or send `Prefer: legacySearch=false`.
- **Photo fetch throttling**: a full page of concurrent `/photos/$value` calls can 429. Lazy-load per visible page, and only mark an id "done" once *definitively* resolved (photo or 404) — never on a transient error, or the photo blanks for the session.
- Batch reads with `$batch` (max 20 sub-requests). Always `$select` only the fields you use.
- Honour `Retry-After` on 429/503; add exponential backoff + jitter otherwise; cap retries.

## Architecture / standards (DevStandards)
- **SPFx / browser → PnPjs**; Azure Node/TS services → `@microsoft/microsoft-graph-client` + `@azure/identity`. SPFx is the only supported custom SharePoint UI model; keep heavy logic in an Azure tier.
- **Microsoft Graph is the default API** for SharePoint list access (less throttling than REST/CSOM).
- **Azure ACS retired 2 Apr 2026** — no client-id+secret add-in auth. Use Entra ID + **managed identity** (`DefaultAzureCredential`) at runtime, interactive `az login` for local scripts, Key Vault references when a secret is unavoidable. Use **`Sites.Selected`**, never tenant-wide `Sites.*.All`.
- **Everything as code**: Bicep for Azure infra (pin AVM versions, `what-if` before apply), PnP templates/CLI for SP schema — idempotent, no ClickOps.

## Coding-standard house rules (for new code)
- `strict: true`; avoid `any` (use `unknown` at boundaries + narrow).
- **No `I`-prefix** on interfaces; types `PascalCase`, values `camelCase`, constants `UPPER_SNAKE`.
- Small single-purpose functions; ≤3 args (else options object); no boolean flag args; no output args.
- Throw `Error` subclasses, never strings; `catch (e: unknown)` then narrow.
- No committed TODOs; no commented-out code; express intent in code over comments.
- React: separate fetch (hook) / logic / presentation; small prop surface; honest `useEffect` deps; never inline `fetch`.
- When editing **existing** files, match the surrounding style (this repo uses `I`-prefixed interfaces) rather than mixing conventions mid-file.
