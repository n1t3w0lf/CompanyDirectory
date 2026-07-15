# UI / UX Strategy — People Directory WebPart

> Keep concise. Records intent + current state of the UI so design decisions stay consistent.

## Principles
- **Fluent UI first** — use Fluent 8 components/tokens; match host SharePoint theme (`supportsThemeVariants: true`).
- **Configurable, not hard-coded** — admins control colours, fonts, visible fields, background via the property pane.
- **Accessible** — WCAG 2.1 AA target: keyboard nav, ARIA labels, focus states, sufficient contrast.
- **Responsive** — card grid reflows; works on mobile + Teams.
- **Fast-perceived** — spinner on load, background photo enrichment, no layout jank.

## Current layout (PeopleDirectory.tsx)
1. **Header** — configurable heading + subtext + live "N people in directory" count.
2. **Sync banner** — shown only when initial sync is required / running.
3. **Search row** — Fluent `SearchBox` + `PrimaryButton` (manual search, no auto-search on type).
4. **Inline filters** — Department / Office Location / City dropdowns + "Clear All Filters".
5. **Letter index** (optional) — A–Z quick filter over the loaded set.
6. **Results** — responsive `userGrid` of `UserCard`s, **client-side paged** (`Pagination.tsx`); empty/loading/error states handled. Empty start shows a Fluent-icon prompt.
7. **Details panel** — `UserDetailsPanel` (Fluent `Panel`, medium) with contact + org sections, click-to-email/call. **Always shows every user property** (the card's visibility toggles do not apply here); empty fields auto-hide.

## Configurable surface (property pane)
Header text/size/colour, subtext, per-field visibility toggles (email, job title, dept, office, phones, city, country, company, employee id), name/jobtitle/properties font size+colour, job-title bold, text ellipsis length, icon size+colour, card background colour/image, web part background colour/image, letter index toggle, profile picture toggle, search button text/colour/size/hover, **People & Pagination** (show people on start, people to show on start, results per page), manual full-sync button.

## Startup & pagination (2026-07-15)
- **Show People on Start** toggle → when OFF, renders a Fluent-icon prompt ("Search for people") instead of loading anyone.
- **People to Show on Start** → count loaded on render (default 30), replaces the old hardcoded 30.
- **Results per Page** → client-side page size (default 20); centred pager below the grid, resets to page 1 on any search/filter/letter change. No "X–Y of Z" range summary (hidden per request). ⚠️ Search fetches ≤100 hits (`manualSearch`), so paging covers up to that cap.
- **Header count removed** — the "N people in directory" line was removed (user edit); the `totalUsers` state + its count fetches were cleaned up.

## Known UX gaps / opportunities (candidate work — confirm before doing)
- **City filter is always empty** — `getCities()` is stubbed to `[]`; dropdown renders but does nothing. Either populate or hide.
- **`AdvancedFilterPanel` component exists but is not rendered** — superseded by inline dropdowns; Country/JobTitle filters have no UI.
- **Server-side paging** — current pager is client-side over the ≤100-hit fetch; large result sets need service-layer paging.
- **"at least 3 characters" copy** vs `MIN_SEARCH_LENGTH = 2` — inconsistent (see discoveries.md).
- **Debug `console.log`** in `webpartContainerStyle` / `UserCard` — remove for production polish.
- No dark-theme-specific handling beyond Fluent defaults; no print view.

## Design references
See `/Users/gulzar/Documents/spfx/designs`, `/UX`, `/UX.zip` at repo root for shared design assets across web parts.
