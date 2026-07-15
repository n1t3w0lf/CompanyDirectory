# Project Working File — People Directory WebPart

> Living log of tasks, decisions, and next steps. Keep concise.
> If this file exceeds **800 lines**, start `PROJECT_WORKING_FILE2.md`.
> Companions: `CODEBASE_CONTEXT.md`, `UI_UX_STRATEGY.md`, `discoveries.md`, `learnings.md`.

## Status (2026-07-15)
- ✅ Full codebase review complete (all `src/`, config, manifest, project docs).
- ✅ DevStandards read + summarised (5 docs).
- ✅ Working docs created: `CODEBASE_CONTEXT.md`, `UI_UX_STRATEGY.md`, `discoveries.md`, `learnings.md`, this file.
- ✅ **Feature set 1 shipped** — property-pane cleanup + startup config + client-side pagination (see below). Build clean (tsc + eslint, Node 18).

## Feature set 1 — Startup & Pagination (2026-07-15)
Delivered:
1. Removed "Property Display Order" pane section + inert `propertyDisplayOrder` prop.
2. Added `showPeopleOnStart` toggle, `initialPeopleCount` slider, `paginationSize` slider ("People & Pagination" group).
3. Startup honours the toggle/count; when OFF shows a Fluent-icon prompt.
4. New `components/Pagination.tsx` — client-side pager; page resets to 1 on search/filter/letter change.
- Rolodex (`showLetterIndex`) left as-is per user.
Files: `PeopleDirectoryWebPart.ts`, `components/IPeopleDirectoryProps.ts`, `components/PeopleDirectory.tsx`, `components/Pagination.tsx` (new).
Build: must use **Node 18** (`nvm use 18`) — Node 22 is rejected by SPFx 1.18 gulp.

### Peer review (adversarial, 2 lenses + verify) — 7 confirmed, all fixed
1. [med] Stale `selectedLetter` hid valid results after search/filter/clear → reset it in those handlers.
2. [low] `currentPage` unclamped vs page-size change → clamp to `safePage` before slicing.
3. [low] `handleStartSync` ignored `showPeopleOnStart=OFF` → mirror `initializeData`.
4. [low] Active page button missing `aria-current="page"` → added.
5. [low] Focus dropped at pager boundary → `allowDisabledFocus` on prev/next.
6. [low] Stale JSDoc "Load first 30 users" → corrected.
(1 finding refuted: theme-token contrast claim — didn't hold up.)

## Feature set 2 — Panel + pager tweaks (2026-07-15)
- **Details panel shows all properties** — `UserDetailsPanel` field toggles forced `true` at the call site; card toggles no longer suppress panel fields.
- **Pager range summary hidden** — removed the "X–Y of Z" text; pager now centred.
- **Header count**: user removed the "N people in directory" line; cleaned up the now-dead `totalUsers` state + `getTotalUserCount` fetches. Build clean (Node 18).

## Feature set 6 — Cache-list schema self-heal (2026-07-15)
Fixed `The property 'PD_Department' does not exist...` 400s during sync. Root cause: `ensureList` short-circuited on an existing list and never added missing columns (interrupted/older creation left the list incomplete); old sync hid it by dropping failures.
- Added `REQUIRED_FIELDS` (single source of truth) + `ensureFields(list)`: reads existing `InternalName`s, adds only missing columns, best-effort re-indexes `INDEXED_FIELDS`. All `ensureList` paths (create / exists / already-exists) now call it. Idempotent, runs once/session.
- Preserves data (adds columns to existing list). Verified: build clean + schema-harness (8 checks: adds missing incl. PD_Department, skips present, correct kinds, indexes 3, read-fail fallback).
- **User action**: rebuild+redeploy, reload once (repairs schema), re-run Full Sync → previously-dropped users write.
- **Indexing added (2026-07-15)**: `INDEXED_FIELDS` extended to the `eq`/`orderBy` columns — `PD_UserPrincipalName, PD_Department, PD_LastVerified, Title, PD_OfficeLocation, PD_City, PD_Country, PD_AccessCount` (applied best-effort by `ensureFields`). Contains-search columns intentionally not indexed (no index benefit). Caveat: index only builds while list <5,000 items; contains-search still scan-bound (see discoveries).
Files: `services/ListService.ts`.

## Feature set 5 — Initial sync speed-up (in-browser) (2026-07-15)
Reworked `SyncService.performInitialSync` + `ListService` writes:
- **Preload existing rows once** (`getExistingUserMap`, keyset-paged on indexed `Id`) → route add-vs-update in memory; removes the per-user existence GET.
- **`bulkUpsertUsers`** writes via PnP `sp.batched()` in chunks of 100 (`@pnp/sp/batching`); de-dupes input by UPN; **honors `Retry-After`** and **retries** throttled items (was: fixed 2.5s/batch + 0.5s/chunk sleeps, and throttled batches were silently **dropped**).
- **No duplicates on re-run**: preloaded map → all existing = updates, none re-added.
- Preserved: on-demand lookup (`manualSearch`→`addOrUpdateUser`) + card-click freshness (`getUserById`→`upsertUser`) untouched.
- `getTotalUserCount` now uses list `ItemCount` (accurate >5000). Removed dead `processBatch`/delay constants. Graph fetch page delay 300→100ms.
- Decisions: **full upsert** each run (no delta), **in-browser** (no Azure Function).
- Verified: Node-18 build clean + CJS mock-harness (15 checks: keyset paging, dedupe/route/chunk, throttle-retry no-drop, re-run=all-updates/zero-adds, first-run=all-adds).
- Expected: ~120K calls + ~3.3h of sleep → ~12 page reads + ~600 batched writes paced by Retry-After. Bounded by SharePoint 5,000 threshold on the preload read (keyset paging handles it).
- **Adversarial review fixes** (verified against PnP `batching.js` source):
  1. 🔴 **Hang**: PnP settles per-op promises only on the `$batch` POST *success* path — so on a batch-level throttle/500 the ops never settle and `await Promise.all(ops)` hung forever. Fixed: on `execute()` throw we no longer await ops — retry the whole chunk (throttle) or abandon it (else); added a cancel check inside the retry loop.
  2. Permanent (non-throttle) per-item failures are now counted in `failed` (were silently uncounted → false "success").
- **Known limitation (documented, not fixed)**: a UPN inserted by another concurrent session *after* the one-shot `existingMap` snapshot could be re-added (duplicate) mid-run. The user's re-run case (all exist → updates) is fully safe; the cross-session race would need `EnforceUniqueValues` (deferred) or a periodic map refresh.
Files: `models/Constants.ts`, `services/ListService.ts`, `services/SyncService.ts`. Harness: 17 checks pass.

## Feature set 4 — Search: all matches + partial/token match (2026-07-15)
- Removed the hardcoded 100 cap: `manualSearch` list branch requests `SEARCH_MAX_RESULTS` (5000); Graph branch pages `@odata.nextLink` (999/page) up to 5000. Client pager pages through them.
- **Partial/"fuzzy" match**: Graph switched `startswith` `$filter` → `$search` (tokenized, mid-name, any order); list already `substringof` (contains), broadened to include given/surname + excludes `_SyncMetadata`.
- Photos now **lazy per visible page** (`enrichPhotos` + a component effect) instead of a blocking bulk fetch; list persistence made non-blocking.
- Verified: Node-18 build clean + a CJS mock-harness (15 checks: `$search` used, paging+5000 cap, list returns all, photo attempt/retry logic).
- Adversarial review fixes: (a) `$search` must be `encodeURIComponent`'d (Graph SDK does no encoding — confirmed in `GraphRequest.js`); (b) sanitize `&`→space to dodge Graph's `%26` `400` bug; (c) photo effect marks only *definitively-resolved* ids so throttled fetches retry instead of blanking.
- Constraints (documented, not built): broad `contains` search bounded by SharePoint's 5,000 threshold; true typo-tolerant fuzzy needs Azure AI Search.
Files: `models/Constants.ts`, `services/GraphService.ts`, `services/PeopleService.ts`, `services/ListService.ts`, `components/PeopleDirectory.tsx`.

## Feature set 3 — Mobile responsive audit (5-inch) (2026-07-15)
Two-agent audit (SCSS + TSX) → fixes; verified with a Playwright CSS harness at **320/360/375/768px**: no horizontal scroll, ≥44px tap targets, rolodex hidden ≤640, desktop intact.
- Card/panel text `word-break`/`overflow-wrap` (kills long-email overflow).
- ≤640px: 44px tap targets via scoped `:global(.ms-Button/.ms-Dropdown-title/.ms-SearchBox)`; filter dropdowns full-width (`.filterItem`); smaller empty icon.
- Rolodex hidden ≤640px (`LetterIndex` display:none) — user decision.
- Pager inner Stack `wrap`.

Adversarial review found 5 real issues (box-size harness had missed them); all fixed + re-verified at 320/375px:
1. Global `:global(.ms-Button){min-width:44}` broke the SearchBox clear "X" → dropped min-width; pager width set per-control via component `styles` media queries.
2. Stuck letter filter on mobile (rolodex hidden) → `getActiveFilterCount()` now counts `selectedLetter` so "Clear All Filters" shows.
3. Dropdown caret misaligned when title grew to 44px → added `.ms-Dropdown-caretDownWrapper` height rule.
4. Name/job-title overflow half-fixed → `overflow-wrap:anywhere` + `min-width:0` on the card name column.
5. Panel close button (portaled, out of scoped CSS) → 44px via Panel `styles.closeButton` media query.
Files: `PeopleDirectory.module.scss`, `LetterIndex.module.scss`, `PeopleDirectory.tsx`, `Pagination.tsx`, `UserCard.tsx`, `UserDetailsPanel.tsx`. Build clean (Node 18).

### ⚠️ Drift incident (resolved)
During the session the `{totalUsers > 0}` "N people in directory" header block was removed from
`PeopleDirectory.tsx` by an external process (not an intended edit; likely a review agent). It exists
in HEAD and is a real feature, so it was **restored**. Lesson: give review/workflow agents read-only
tools. Verify final `git diff HEAD` before committing.

## Decisions log
- **2026-07-15** — Established the mandated working-doc set (this file + 4 companions).
- **2026-07-15** — Standards from `/Users/gulzar/Documents/spfx/DevStandards/` are the gate for any future change (see checklist below).
- **2026-07-15** — Pagination is **client-side** over the loaded result set (simple, no service change); flagged the ≤100-hit search fetch cap. Server-side paging deferred.
- **2026-07-15** — Kept `react/jsx-no-bind` clean by using bound `useCallback`/sub-component handlers (matches `LetterIndex` pattern), not inline arrows.

## Task backlog (not started — awaiting go-ahead)
Ranked; details in `discoveries.md`.

1. **Correctness bugs (small, safe, high-value)**
   - `_SyncMetadata` row leaks into `getInitialUsers`/`searchUsers` results.
   - id-vs-UPN mismatch in `PeopleService.getUserById` list fallback.
   - Search-length copy says "3" but `MIN_SEARCH_LENGTH = 2`.
2. **Scalability rework (architectural)** — 5,000 list-view threshold vs `LIST_MAX_ITEMS = 65000`: index columns, keyset paging, real count query, `$batch` photos.
3. **Complete inert features** — City/Country/JobTitle filters (stubbed to `[]`). (`propertyDisplayOrder` removed; pagination now shipped in place of `react-window` for large sets.)
4. **Cleanup** — remove debug `console.log`s; structured logging / App Insights hook.

## Standards checklist (gate for any future change)
- SPFx/browser data access → **PnPjs**; keep heavy logic out of the browser.
- Graph reads: `$select` + `$filter` on an **indexed** column + paging (`$top`+`nextLink`) + `$batch` (≤20) + ETag `If-Match` on writes.
- Respect the **5,000 list-view threshold**; never scan an unbounded list.
- New TS code: `strict`, avoid `any`, **no `I`-prefix** on interfaces, small functions, no flag args, no committed TODOs. (Existing code uses `I`-prefix + `PascalCase` interfaces — match local style when editing existing files; apply house rules to new files.)
- No secrets anywhere; Entra ID + managed identity / `Sites.Selected` for any Azure-side work.
- WCAG 2.1 AA; Fluent UI + theme variants; mobile-first.
- No debug `console.log` in production paths; use `ErrorHandler` + structured logging.
- Peer-review own changes and ensure zero compile/lint errors before calling a task done.

## Next steps
1. ✅ Finish docs (this file + `learnings.md`).
2. ⏭️ Wait for the user to pick a backlog item, then plan it in detail before coding.
