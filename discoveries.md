# Discoveries — People Directory WebPart

> Things found while reading the code. Facts + potential issues. Newest at top.
> Severity: 🔴 bug/risk · 🟠 correctness/perf · 🟡 polish/cleanup.

## 2026-07-15 — Initial full code review

### Correctness / bugs
- 🔴 **`_SyncMetadata` row leaks into user lists.** `getInitialUsers` → `getPaginatedUsers` and `ListService.searchUsers` do **not** exclude `Title eq '_SyncMetadata'`. Only `getFilteredUsers` excludes it. So the metadata row can appear as a fake "person" card. (`ListService.ts:200`, `:320`)
- 🔴 **Fallback lookup passes id where UPN expected.** `PeopleService.getUserById` catch-block calls `listService.getUser(userId)`, but `ListService.getUser` filters on `PD_UserPrincipalName`. When called with a Graph object id, the fallback silently returns null. (`PeopleService.ts:116`, `ListService.ts:169`)
- 🟠 **Search-length copy mismatch.** `Constants.MIN_SEARCH_LENGTH = 2`, but UI error text says "at least 3 characters". (`PeopleDirectory.tsx:167-168`)

### Scale / performance
- 🔴 **5,000 list-view threshold vs `LIST_MAX_ITEMS = 65000`.** SharePoint enforces a 5,000-item view threshold. Filtering/ordering on non-indexed columns, and `.skip()` pagination over large lists, will throw or degrade badly well before 65K. Ordering by `Title` (not confirmed indexed) is a risk. Design assumes list holds the whole org.
- 🟠 **Counting by fetching rows.** `getTotalUserCount` pulls up to 5,000 items just to `.length` them — expensive and capped at 5,000 (wrong count for bigger dirs). `upsertUser` does the same to check capacity. Prefer `list.select('ItemCount')` / a count query.
- 🟠 **Per-user photo calls.** `enrichUsersWithPhotos` issues one Graph call per user (30+ on initial load). Not batched via `$batch`; no caching of "no photo".
- 🟠 **`getDepartments`/`getOfficeLocations` on Graph** paginate the entire directory (`$top=999` loop) — very heavy for 60K; only used as fallback but worth guarding.

### Stubs / inert features
- 🟠 `getCities`, `getCountries`, `getJobTitles` in `PeopleService` return `[]` (TODO stubs). City dropdown is therefore always empty.
- ✅ ~~`propertyDisplayOrder` inert~~ — **removed 2026-07-15** (feature set 1).
- 🟡 `AdvancedFilterPanel.tsx` not rendered anywhere; superseded by inline dropdowns.
- 🟡 `react-window` installed but virtual scrolling not implemented (client-side pagination added instead, 2026-07-15).

### Cleanup
- 🟡 Debug `console.log` left in `PeopleDirectory` (`webpartContainerStyle`) and `UserCard` (`cardStyle`). Standards discourage console logging in production paths.
- 🟡 Broad `console.error`/`console.log` throughout services instead of a structured logger; `ErrorHandler.logError` has a TODO for App Insights.

### Confirmed-good patterns (don't "fix")
- ✅ OData string escaping (`replace(/'/g, "''")`) applied to Graph `$filter` and SP `filter` inputs.
- ✅ `PD_` field prefix avoids SharePoint reserved-name collisions.
- ✅ `parseNextLink` strips `/v1.0`|`/beta` because MSGraphClientV3 re-adds the version.
- ✅ Cache writes wrapped so failures don't break UX; `Promise.allSettled` for batch enrich.
