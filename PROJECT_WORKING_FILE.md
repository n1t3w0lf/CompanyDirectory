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
