# Codebase Context — People Directory WebPart

> Purpose: fast orientation for anyone (human or AI) picking up this project.
> Keep concise. Update when architecture changes.

## What it is
Enterprise SPFx web part that searches/browses people from **Microsoft Graph**, cached in a **SharePoint list** and **IndexedDB**. Targets large orgs (stated 50K–60K users).

## Stack
- SPFx **1.18.2**, React **17.0.1**, TypeScript **4.7.4**
- Fluent UI React **8.x**, PnPjs **3.22+**, react-window (installed, not yet used)
- Node 16.13+ / 18.17+

## Layers
| Layer | Files | Role |
|-------|-------|------|
| Web part | `PeopleDirectoryWebPart.ts` | onInit wires services, property pane, renders React |
| UI | `components/*.tsx` | `PeopleDirectory` (container), `UserCard`, `UserDetailsPanel`, `LetterIndex`, `SyncStatusBanner`, `Pagination` (client-side pager), `AdvancedFilterPanel` (unused) |
| Services | `services/*.ts` | `GraphService`, `ListService`, `PeopleService` (orchestrator), `SyncService` (bulk import) |
| Utils | `utils/*.ts` | `CacheHelper` (IndexedDB), `ErrorHandler` |
| Models | `models/*.ts` | `IUserProfile` + interfaces, `Constants` |

## Data flow (current, as-built)
1. `onInit` → create MSGraphClientV3 + PnP `SPFI`, build services.
2. `PeopleDirectory` mounts → `peopleService.initialize()` (IndexedDB + `ensureList`).
3. If no sync metadata → show **SyncStatusBanner** prompting initial sync.
   Else → `getInitialUsers(30)` from the **SharePoint list** (list is now primary read source).
4. Manual search: list first → fall back to Graph → write new hits back to list.
5. Card click → `getUserById(id, forceRefresh)` re-verifies against Graph.

## SharePoint list — `PeopleDirectoryCache`
- Custom fields prefixed **`PD_`** (avoids SP reserved names). Title = displayName.
- Indexed: `PD_UserPrincipalName`, `PD_Department`, `PD_LastVerified`.
- A special row `Title = '_SyncMetadata'` stores last-sync info (reuses PD_ columns).
- `Constants.LIST_MAX_ITEMS = 65000` — **note the 5,000 list-view threshold tension** (see discoveries.md).

## Graph
- Permissions requested: `User.ReadBasic.All`, `User.Read.All` (config/package-solution.json).
- `GRAPH_SELECT_FIELDS` fixed set; photos via `/photos/96x96/$value` → base64 data URL.
- Pagination via `@odata.nextLink`; `parseNextLink` strips `/v1.0` (client re-adds it).
- Advanced queries send `ConsistencyLevel: eventual`.

## Config / identity
- Solution id `8f4a7a9e-…`, version `1.0.0.4`; web part id `7c4a8b9f-…`.
- Hosts: SharePointWebPart, TeamsPersonalApp, TeamsTab. `supportsThemeVariants: true`.

## Build / run
- `npm run serve` (workbench), `npm run package` (ship .sppkg).
- Output: `sharepoint/solution/people-directory-webpart.sppkg`.

## Dev standards
Repo-wide standards live in `/Users/gulzar/Documents/spfx/DevStandards/`. Follow them for all changes (see PROJECT_WORKING_FILE.md → Standards checklist).

## Companion docs
`ARCHITECTURE.md` (design intent — note: some of it is aspirational vs as-built), `README.md`, `DEPLOYMENT.md`, `QUICKSTART.md`, `UI_UX_STRATEGY.md`, `PROJECT_WORKING_FILE.md`, `learnings.md`, `discoveries.md`.
