# Plan — `GraphMyDayDataService` (live Microsoft Graph data)

**Status:** approved (all five open questions answered) — implementation in progress.

Approved answers: (1) async contract yes; (2) all four Graph permissions yes;
(3) Settings-drawer toggle yes; (4) configurable news site yes; (5) weather
comes from a live weather API rather than staying mocked (see §3 D8).

## 1. Goal

Add a `GraphMyDayDataService` that fulfils the existing `IMyDayDataService`
contract using the SPFx-brokered `MSGraphClientV3`, pulling **meetings, tasks,
mail and news** from Microsoft Graph, and add a **configuration switch** so the
component can run against either the mock provider or the live Graph provider.

Non-goals for this change: replacing the mock data (it stays, and stays the
default), changing the visual design, or introducing any new UI panels.

## 2. Current state

| Piece | Today |
| --- | --- |
| [IMyDayDataService.ts](samples/my-day/src/copilotComponents/myDay/services/IMyDayDataService.ts) | `getMyDay(now?: Date): IMyDayData` — **synchronous** |
| [MockMyDayDataService.ts](samples/my-day/src/copilotComponents/myDay/services/MockMyDayDataService.ts) | Resolves bundled mock arrays, maps, sorts |
| [mappers.ts](samples/my-day/src/copilotComponents/myDay/services/mappers.ts) | Graph shape → view model. Already written to be shared |
| [models/graph.ts](samples/my-day/src/copilotComponents/myDay/models/graph.ts) | Graph-shaped subsets (`IGraphEvent`, `IGraphTodoTask`, `IGraphSitePage`, `IGraphMessage`) |
| [MyDayInline.tsx](samples/my-day/src/copilotComponents/myDay/components/MyDayInline.tsx) / [MyDayFullscreen.tsx](samples/my-day/src/copilotComponents/myDay/components/MyDayFullscreen.tsx) | Each constructs `new MockMyDayDataService()` at module scope and calls it inside `React.useMemo` |
| [CurrentUserService.ts](samples/my-day/src/copilotComponents/myDay/services/CurrentUserService.ts) | Resolves the signed-in user from the SPFx page context (no Graph call) |
| [config/package-solution.json](samples/my-day/config/package-solution.json) | No `webApiPermissionRequests` |

The mappers and Graph models are already the right shape, so **no mapper
rewrite is expected** — the Graph service can feed real payloads straight
through `mapMeeting` / `mapTask` / `mapNews` / `mapMail`.

## 3. Key design decisions (please confirm)

### D1 — The service contract becomes async

`getMyDay` must return `Promise<IMyDayData>`. This is the single genuinely
breaking change and it ripples into both views.

```ts
export interface IMyDayDataService {
  getMyDay(now?: Date): Promise<IMyDayData>;
}
```

`MockMyDayDataService` keeps its synchronous internals and simply returns a
resolved promise, so the mock stays instant and deterministic.

### D2 — Views gain a loading / error state

`React.useMemo` is replaced with a small `useMyDayData(service, now)` hook
(`services/useMyDayData.ts`) returning `{ data, loading, error }`. Views render:

- **loading** — the existing panel skeleton/shell with a Fluent `Spinner`
  (fullscreen) or a compact shimmer (inline);
- **error** — never a hard failure; see D5.

The mock path resolves in a microtask, so the demo experience is unchanged.

### D3 — Configuration: a component property plus a factory

Add to [MyDayCopilotComponentProperties.ts](samples/my-day/src/copilotComponents/myDay/MyDayCopilotComponentProperties.ts):

```ts
dataSource: z.enum(['mock', 'graph']).default('mock')
  .describe('Where My Day reads its data from.')
```

and a factory, `services/createMyDayDataService.ts`, that is the **only** file
importing `@microsoft/sp-http` — keeping both services free of SPFx types and
unit-testable:

```ts
export const createMyDayDataService = (
  context: CopilotComponentContext,
  source: MyDayDataSource
): IMyDayDataService =>
  source === 'graph'
    ? new GraphMyDayDataService(context.msGraphClientFactory, graphOptions)
    : new MockMyDayDataService();
```

`MyDayCopilotComponent` builds the service once and passes it down through
`MyDayApp` → `MyDayInline` / `MyDayFullscreen` as a prop (removing the two
module-scope `new MockMyDayDataService()` instances).

**Default stays `'mock'`** so the sample deploys and demos with zero admin
consent, matching the README promise.

**Approved add-on:** a "Data source" toggle in the fullscreen Settings drawer,
persisted via the existing
[settings.ts](samples/my-day/src/copilotComponents/myDay/utils/settings.ts)
`sessionStorage` hook, so a demo can flip mock ↔ Graph live. The `dataSource`
property provides the initial value; the session setting wins once the user
changes it.

### D4 — Graph endpoints per section

All calls are **delegated** (signed-in user) via `MSGraphClientV3`, issued in
parallel with `Promise.allSettled`.

| Section | Endpoint | Notes |
| --- | --- | --- |
| Meetings | `GET /me/calendarView?startDateTime={todayStart}&endDateTime={todayEnd}` | `$orderby=start/dateTime`, `$top=25`, `Prefer: outlook.timezone="<user tz>"` so `start.dateTime` is local — matches what `mapMeeting` assumes |
| Tasks | `GET /me/todo/lists` → pick `wellknownListName === 'defaultList'`, then `GET /me/todo/lists/{id}/tasks` | `$filter=status ne 'completed'`, `$top=20`. Graph returns `importance` and `dueDateTime` already |
| Mail | `GET /me/mailFolders/inbox/messages` | `$select` the mapped fields, `$filter=isRead eq false or importance eq 'high'`, `$orderby=receivedDateTime desc`, `$top=10` |
| News | `GET /sites/{siteId}/pages/microsoft.graph.sitePage` (v1.0) | Site is configurable (see below); filter `promotionKind eq 'newsPost'`, `$orderby=createdDateTime desc`, `$top=10` |
| User | unchanged — page context | `CurrentUserService` already resolves name + `userphoto.aspx` photo with no Graph call or permission |
| Weather | live weather API — see D8 | Not a Graph resource |
| Quick actions | unchanged — static | No Graph source exists; the Graph service reuses the same static tiles |

**News source configuration.** `sitePage` requires a site id. Proposal: a second
optional component property `newsSiteUrl` (e.g. the intranet home site). When
empty, resolve the tenant home site via `GET /sites/root` and read its pages;
if that yields nothing, the news panel falls back to the mock news (D5).

**Sender / author photos.** `IMailItem.senderPhotoUrl` and the news byline photo
are populated with the SharePoint `userphoto.aspx` URL built by the existing
`buildUserPhotoUrl` helper — no extra Graph round-trip, no `User.ReadBasic.All`
requirement, and a built-in silhouette fallback.

### D5 — Graceful degradation, never a blank dashboard

Each section resolves independently. A rejected section:

1. logs a single console warning with the section name and Graph error code;
2. yields an **empty array** for that section, so the panel shows its existing
   empty state;
3. surfaces a non-blocking `IMyDayDataIssue[]` on the result, which the
   fullscreen view renders as a small inline `MessageBar` (e.g. "News is
   unavailable right now").

If **every** section fails (no consent, offline), the factory result falls back
wholesale to `MockMyDayDataService` so the component still renders.

### D6 — Permissions and admin consent

Add to `config/package-solution.json`:

```jsonc
"webApiPermissionRequests": [
  { "resource": "Microsoft Graph", "scope": "Calendars.Read" },
  { "resource": "Microsoft Graph", "scope": "Tasks.Read" },
  { "resource": "Microsoft Graph", "scope": "Mail.Read" },
  { "resource": "Microsoft Graph", "scope": "Sites.Read.All" }
]
```

These are requested but **only consumed when `dataSource: 'graph'`**. Because
they are broad, the README will call out clearly that a tenant admin must
approve them in the SharePoint admin centre before the Graph mode works, and
that the default mock mode needs none of them.

### D8 — Weather from a live API (Open-Meteo)

Weather is not a Graph resource, so it gets its own small provider,
`services/OpenMeteoWeatherService.ts`, used only in `graph` mode:

- **Open-Meteo** — free, no API key, no sign-up, CORS-enabled. Choosing a
  keyless API keeps the sample deployable by anyone, which an
  OpenWeatherMap/AccuWeather key would not.
- Geocode the `city` / `country` already captured in the Settings drawer via
  `https://geocoding-api.open-meteo.com/v1/search?name={city}&count=1`.
- Current conditions via
  `https://api.open-meteo.com/v1/forecast?latitude=..&longitude=..&current=temperature_2m,weather_code`
  and air quality via
  `https://air-quality-api.open-meteo.com/v1/air-quality?...&current=european_aqi`.
- WMO `weather_code` is mapped to the existing `IWeather.condition` string, and
  the European AQI to `airQualityIndex` / `airQualityLabel`.
- Called with SPFx `HttpClient` (anonymous, no credentials sent). Failure or a
  city that will not geocode degrades to the mock weather — the hero never
  breaks.

> Deployment note for the README: the Open-Meteo domains must be reachable from
> the browser. No tenant configuration is normally required, but locked-down
> networks may block them, in which case weather silently falls back to mock.

### D7 — Performance and resilience details

- One `MSGraphClientV3` instance resolved once per service, reused for all calls.
- All four section queries fire concurrently; total latency ≈ slowest call.
- `$select` on every request to keep payloads small.
- Results cached in-memory per service instance for 60s so the inline →
  fullscreen transition does not re-query.
- 429/503 responses are **not** retried in this pass (single attempt, then
  degrade) — a `Retry-After` aware retry can follow if you want it.

## 4. Files to add / change

**New**

- `services/GraphMyDayDataService.ts` — the implementation
- `services/graphQueries.ts` — one small function per section, each returning the
  Graph-shaped models from `models/graph.ts`
- `services/createMyDayDataService.ts` — the factory (only SPFx-aware file)
- `services/useMyDayData.ts` — loading/error hook
- `services/OpenMeteoWeatherService.ts` — keyless live weather (D8)
- `components/fullscreen/settings/…` — "Data source" toggle in the drawer

**Changed**

- `services/IMyDayDataService.ts` — async signature, `issues` on the result
- `services/MockMyDayDataService.ts` — returns a resolved promise
- `models/myDay.ts` — optional `issues?: IMyDayDataIssue[]` on `IMyDayData`
- `MyDayCopilotComponentProperties.ts` — `dataSource`, `newsSiteUrl`
- `MyDayCopilotComponent.ts` — build the service, pass it down
- `components/MyDayApp.tsx` / `MyDayInline.tsx` / `MyDayFullscreen.tsx` — accept
  the service prop, use the hook, render loading / issue states
- `config/package-solution.json` — `webApiPermissionRequests`
- `copilot/declarativeAgent.json` + `instruction.txt` — only if the new
  properties should be settable by the agent (recommend: **no**, keep
  `dataSource` a deploy-time decision, so no change here)
- `README.md` — new "Live Microsoft Graph data" section: permissions, consent
  steps, how to switch, and what stays mocked

**Unchanged**

- `mappers.ts`, `models/graph.ts`, all presentational components, mock data

## 5. Verification

- Existing lint/build (`npm run build`) clean.
- Manual: `dataSource: 'mock'` — behaviour identical to today.
- Manual: `dataSource: 'graph'` in a consented tenant — meetings, tasks and mail
  match Outlook/To Do; news matches the configured site.
- Manual: `dataSource: 'graph'` **without** consent — dashboard still renders,
  with empty panels and the informational message bar.
- Unit-level sanity: `GraphMyDayDataService` against a stubbed client (the
  service takes a factory, so no SPFx runtime is needed).

## 6. Resolved questions

1. **Async contract** — approved.
2. **Permissions** — all four approved (`Calendars.Read`, `Tasks.Read`,
   `Mail.Read`, `Sites.Read.All`).
3. **Settings-drawer toggle** — approved, session-persisted.
4. **News source** — configurable `newsSiteUrl` property with `/sites/root`
   fallback.
5. **Weather** — live via Open-Meteo (D8), mock as the fallback.
