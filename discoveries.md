# Discoveries — People Directory WebPart

> Things found while reading the code. Facts + potential issues. Newest at top.
> Severity: 🔴 bug/risk · 🟠 correctness/perf · 🟡 polish/cleanup.

## 2026-07-15 — Initial full code review

### Correctness / bugs
- 🔴 **`_SyncMetadata` row leaks into user lists.** `getInitialUsers` → `getPaginatedUsers` and `ListService.searchUsers` do **not** exclude `Title eq '_SyncMetadata'`. Only `getFilteredUsers` excludes it. So the metadata row can appear as a fake "person" card. (`ListService.ts:200`, `:320`)
- 🔴 **Fallback lookup passes id where UPN expected.** `PeopleService.getUserById` catch-block calls `listService.getUser(userId)`, but `ListService.getUser` filters on `PD_UserPrincipalName`. When called with a Graph object id, the fallback silently returns null. (`PeopleService.ts:116`, `ListService.ts:169`)
- ✅ ~~**Search-length copy mismatch** ("3 characters" vs `MIN_SEARCH_LENGTH = 2`)~~ — fixed 2026-07-15 (now interpolates the constant).

### Scale / performance
- 🟠 **Indexing (partly resolved 2026-07-15).** `INDEXED_FIELDS` now covers the `eq`/`orderBy` query columns: `PD_UserPrincipalName`, `PD_Department`, `PD_LastVerified`, `Title`, `PD_OfficeLocation`, `PD_City`, `PD_Country`, `PD_AccessCount` (applied best-effort by `ensureFields`). Deliberately NOT indexed: the `substringof` (contains) search columns (`PD_Email`/`PD_JobTitle`/`PD_GivenName`/`PD_Surname`) — SharePoint can't use an index for contains. Remaining caveats: (a) SharePoint only builds an index while the list is **<5,000 items** → index before a big sync; on an already-large list the index calls fail (logged, non-fatal); (b) contains-search is still scan/5,000-bound → true 60K search needs a search index (Azure AI Search / SP Search). Max 20 indexed cols.
- ✅ ~~**Cache-list schema drift** — `ensureList` never added missing columns to an existing list → `PD_Department does not exist` 400s on every write~~ — fixed 2026-07-15 (feature set 6: `ensureFields` self-heal).
- 🔴 **5,000 list-view threshold vs `LIST_MAX_ITEMS = 65000`.** SharePoint enforces a 5,000-item view threshold. Filtering/ordering on non-indexed columns, and `.skip()` pagination over large lists, will throw or degrade badly well before 65K. (Sync's bulk read now uses keyset paging on indexed `Id`; search/`getPaginatedUsers` still `.skip()`/threshold-bound.)
- ✅ ~~**Slow initial sync** — 2 REST calls/user + fixed 2.5s/batch + 0.5s/chunk sleeps (~3.3h for 60K); throttled batches silently **dropped** users.~~ — fixed 2026-07-15 (feature set 5: preload map + `$batch` + Retry-After retry).
- ✅ ~~**Counting by fetching rows** (`getTotalUserCount` `.top(5000).length`, capped/wrong >5000)~~ — fixed 2026-07-15 (reads list `ItemCount`). Note `upsertUser`'s dead capacity-count query remains (only the interactive insert path).
- 🟠 **Per-user photo calls.** `enrichUsersWithPhotos` issues one Graph call per user (30+ on initial load). Not batched via `$batch`; no caching of "no photo".
- 🟠 **`getDepartments`/`getOfficeLocations` on Graph** paginate the entire directory (`$top=999` loop) — very heavy for 60K; only used as fallback but worth guarding.

### Stubs / inert features
- 🟠 `getCities`, `getCountries`, `getJobTitles` in `PeopleService` return `[]` (TODO stubs). City dropdown is therefore always empty.
- ✅ ~~`propertyDisplayOrder` inert~~ — **removed 2026-07-15** (feature set 1).
- 🟡 `AdvancedFilterPanel.tsx` not rendered anywhere; superseded by inline dropdowns.
- 🟡 `react-window` installed but virtual scrolling not implemented (client-side pagination added instead, 2026-07-15).

### Mobile responsive (audited + fixed 2026-07-15)
- ✅ ~~`.userInfo`/`.userJobTitle` no word-break → long emails overflow~~ — fixed (word-break/overflow-wrap + min-width:0).
- ✅ ~~Pager inner row couldn't wrap → horizontal scroll~~ — fixed (`wrap`).
- ✅ ~~Sub-44px tap targets (rolodex, pager, search, dropdowns)~~ — fixed (44px `:global` rules ≤640; rolodex hidden ≤640).
- ✅ ~~Filter dropdowns fixed 200px on mobile~~ — fixed (`.filterItem` full-width ≤640).
- Note: media queries are viewport-based; a narrow section column on a wide screen keeps desktop rules.

### Cleanup
- 🟡 Debug `console.log` left in `PeopleDirectory` (`webpartContainerStyle`) and `UserCard` (`cardStyle`). Standards discourage console logging in production paths.
- 🟡 Broad `console.error`/`console.log` throughout services instead of a structured logger; `ErrorHandler.logError` has a TODO for App Insights.

### Confirmed-good patterns (don't "fix")
- ✅ OData string escaping (`replace(/'/g, "''")`) applied to Graph `$filter` and SP `filter` inputs.
- ✅ `PD_` field prefix avoids SharePoint reserved-name collisions.
- ✅ `parseNextLink` strips `/v1.0`|`/beta` because MSGraphClientV3 re-adds the version.
- ✅ Cache writes wrapped so failures don't break UX; `Promise.allSettled` for batch enrich.
