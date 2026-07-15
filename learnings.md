# Learnings — People Directory WebPart

> General, reusable knowledge (SPFx / Graph / SharePoint / standards).
> Codebase-specific findings + issues live in `discoveries.md`. Newest at top.

## Build / tooling
- **SPFx 1.18 gulp needs Node 16/18** — Node 22 is rejected outright. Use `nvm use 18` (available) before any `gulp build/bundle/serve`. `tsc` alone runs on 22 but skips lint + scss codegen.
- The generated `*.module.scss.ts` lists the CSS-module class names TS can reference; reuse existing classes or keep new controls self-contained (Fluent + inline) to avoid regenerating it.
- Fluent 8 v8 has **no Pagination** in `@fluentui/react` core — build a small pager from `DefaultButton`/`IconButton`.
- This repo's ESLint enforces **`react/jsx-no-bind`** — no inline arrows in JSX; use `useCallback`/sub-components (see `LetterIndex`, `Pagination`).
- **A11y is a build gate here** — active nav item needs `aria-current`; keep boundary-disabled buttons focusable with `allowDisabledFocus`.

## Workflow / process
- Give review/analysis subagents **read-only** tools. A review workflow once mutated a source file (removed a feature block) — always run `git diff HEAD` and rebuild before trusting a post-review tree.

## SharePoint / PnPjs
- **5,000 list-view threshold is fixed** (per-query, not storage) and cannot be raised. Index every column you filter/sort on — at creation, while the list is small. Max 20 indexed columns. To read a whole large list: keyset-page on the indexed key (`Id gt <lastId>` ordered by `Id`) then filter/sort in memory. A bare `.top(5000)` is silent truncation, not a safe cap.
- **`PD_` field prefix** avoids SharePoint reserved-name collisions on custom columns.
- Code should bind to a column's **internal name**, never its display name.
- Prefer counting via list metadata (`ItemCount`) over fetching rows and taking `.length`.

## Microsoft Graph
- **`MSGraphClientV3` auto-prepends the `/v1.0` (or `/beta`) version prefix.** When reusing an `@odata.nextLink`, strip the version segment first or the path doubles up (see `GraphService.parseNextLink`).
- **`ConsistencyLevel: eventual`** header is required for `$count` and advanced query params.
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
