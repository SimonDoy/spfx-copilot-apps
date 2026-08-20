import type { MSGraphClientV3 } from '@microsoft/sp-http';

import type {
  IGraphEvent,
  IGraphMessage,
  IGraphSitePage,
  IGraphTodoTask
} from '../models/graph';

/**
 * One function per My Day section, each returning the Graph-shaped models
 * declared in `models/graph.ts` so the existing mappers can be reused verbatim.
 *
 * Every query is delegated (runs as the signed-in user) through the SPFx
 * brokered `MSGraphClientV3` — no token handling here. Errors are allowed to
 * propagate; `GraphMyDayDataService` decides how to degrade.
 */

interface IGraphCollection<T> {
  value?: T[];
}

/** The IANA time zone of the browser, used to ask Graph for local event times. */
export const resolveTimeZone = (): string => {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  } catch {
    return 'UTC';
  }
};

const startOfDay = (now: Date): Date =>
  new Date(now.getFullYear(), now.getMonth(), now.getDate());

const endOfDay = (now: Date): Date =>
  new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);

/** Local wall-clock ISO string (no `Z`) — what `calendarView` expects with a `Prefer` time zone. */
const toLocalIso = (d: Date): string => {
  const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` +
    `T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
  );
};

/** Today's calendar, in the user's own time zone. */
export const fetchTodaysEvents = async (
  client: MSGraphClientV3,
  now: Date
): Promise<IGraphEvent[]> => {
  const response: IGraphCollection<IGraphEvent> = await client
    .api('/me/calendarView')
    .header('Prefer', `outlook.timezone="${resolveTimeZone()}"`)
    .query({
      startDateTime: toLocalIso(startOfDay(now)),
      endDateTime: toLocalIso(endOfDay(now))
    })
    .select(
      'id,subject,start,end,location,isOnlineMeeting,onlineMeeting,onlineMeetingProvider,organizer,importance,showAs,isAllDay,webLink'
    )
    .orderby('start/dateTime')
    .top(25)
    .get();

  return response.value || [];
};

interface IGraphTodoList {
  id: string;
  displayName?: string;
  wellknownListName?: string;
}

/** Open tasks from the user's default To Do list. */
export const fetchOpenTasks = async (client: MSGraphClientV3): Promise<IGraphTodoTask[]> => {
  const lists: IGraphCollection<IGraphTodoList> = await client.api('/me/todo/lists').get();
  const all = lists.value || [];
  const defaultList =
    all.filter((l) => l.wellknownListName === 'defaultList')[0] || all[0];

  if (!defaultList) {
    return [];
  }

  const response: IGraphCollection<IGraphTodoTask> = await client
    .api(`/me/todo/lists/${defaultList.id}/tasks`)
    .filter("status ne 'completed'")
    .top(20)
    .get();

  return response.value || [];
};

/**
 * Recent inbox mail. Deliberately unfiltered server-side: Outlook rejects an
 * `$orderby` that does not lead with the filtered property, so the "important"
 * narrowing (unread or high importance) happens client-side.
 */
export const fetchImportantMail = async (
  client: MSGraphClientV3,
  top: number = 10
): Promise<IGraphMessage[]> => {
  const response: IGraphCollection<IGraphMessage> = await client
    .api('/me/mailFolders/inbox/messages')
    .select(
      'id,subject,bodyPreview,from,receivedDateTime,isRead,importance,hasAttachments,flag,webLink'
    )
    .orderby('receivedDateTime desc')
    .top(50)
    .get();

  const messages = (response.value || []).filter((m) => !!m.from);
  const important = messages.filter((m) => !m.isRead || m.importance === 'high');

  return (important.length > 0 ? important : messages).slice(0, top);
};

/**
 * Resolves a SharePoint site URL to a Graph site id. Falls back to the tenant
 * root site when no URL is configured.
 */
export const resolveSiteId = async (
  client: MSGraphClientV3,
  siteUrl?: string
): Promise<string | undefined> => {
  let path = '/sites/root';

  if (siteUrl) {
    try {
      const url = new URL(siteUrl);
      const serverRelative = url.pathname.replace(/\/$/, '');
      path = `/sites/${url.hostname}:${serverRelative}`;
    } catch {
      // Malformed URL — fall through to the tenant root site.
    }
  }

  const site: { id?: string } = await client.api(path).select('id').get();
  return site.id;
};

interface IGraphSitePageResponse extends IGraphSitePage {
  createdBy?: { user?: { displayName?: string; email?: string } };
}

/**
 * News posts from a SharePoint site. `promotionKind` is filtered client-side
 * because it is not a supported `$filter` property on `sitePage`.
 */
export const fetchNews = async (
  client: MSGraphClientV3,
  siteId: string,
  buildAuthorPhotoUrl: (email: string) => string | undefined,
  top: number = 10
): Promise<IGraphSitePage[]> => {
  const response: IGraphCollection<IGraphSitePageResponse> = await client
    .api(`/sites/${siteId}/pages/microsoft.graph.sitePage`)
    .select('id,name,title,webUrl,description,thumbnailWebUrl,createdDateTime,promotionKind')
    .expand('createdBy')
    .orderby('createdDateTime desc')
    .top(40)
    .get();

  return (response.value || [])
    .filter((p) => p.promotionKind === 'newsPost')
    .slice(0, top)
    .map((p) => {
      const author = p.createdBy?.user;
      return {
        ...p,
        author: author?.displayName
          ? {
              displayName: author.displayName,
              email: author.email,
              photoUrl: author.email ? buildAuthorPhotoUrl(author.email) : undefined
            }
          : undefined
      };
    });
};
