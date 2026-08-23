import type { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

import { mockQuickActions } from '../mockData';
import type { IMyDayData, IMyDayDataIssue, IUser, MyDaySection } from '../models/myDay';
import { fetchTodaysEvents } from './graphQueries';
import type { IMyDayDataService, IMyDayQueryOptions } from './IMyDayDataService';
import { mapMeeting } from './mappers';
import type { IWeatherService } from './OpenMeteoWeatherService';

export interface IGraphMyDayDataServiceOptions {
  /** Signed-in user, already resolved from the SPFx page context. */
  currentUser: IUser;
  /** Absolute URL of the current web, used to build `userphoto.aspx` URLs. */
  webAbsoluteUrl: string;
  /** SharePoint site to read news from. Defaults to the tenant root site. */
  newsSiteUrl?: string;
  /** Live weather provider. Omitted in environments with no outbound access. */
  weatherService?: IWeatherService;
}

/** How long a resolved result is reused, so inline → fullscreen does not re-query. */
const CACHE_TTL_MS = 60_000;

/**
 * How long to wait for a Graph client (i.e. a signed-in token) before giving
 * up. The underlying MSAL/AadHttpClient silent flow can hang far longer than
 * this (e.g. `BrowserAuthError: monitor_window_timeout` inside a nested
 * iframe), which would otherwise leave the dashboard on the loading screen
 * indefinitely.
 */
const CLIENT_TIMEOUT_MS = 8_000;

/** Rejects with `message` if `promise` has not settled within `ms`. */
const withTimeout = <T>(promise: Promise<T>, ms: number, message: string): Promise<T> =>
  new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), ms);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error: unknown) => {
        clearTimeout(timer);
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    );
  });

const SECTION_LABELS: { [K in MyDaySection]: string } = {
  meetings: 'Your calendar',
  tasks: 'Your tasks',
  mail: 'Your mail',
  news: 'News',
  weather: 'Weather'
};

/**
 * Live implementation of {@link IMyDayDataService} backed by Microsoft Graph.
 *
 * All four Graph sections are queried concurrently under the signed-in user's
 * own identity via the SPFx-brokered client. Sections fail independently: a
 * rejected call yields an empty section plus a non-blocking issue, so the
 * dashboard always renders.
 */
export class GraphMyDayDataService implements IMyDayDataService {
  private readonly clientFactory: MSGraphClientFactory;
  private readonly options: IGraphMyDayDataServiceOptions;
  private cache: { key: string; at: number; data: Promise<IMyDayData> } | undefined;

  public constructor(
    clientFactory: MSGraphClientFactory,
    options: IGraphMyDayDataServiceOptions
  ) {
    this.clientFactory = clientFactory;
    this.options = options;
     console.log('GraphMyDayDataService constructed', { options });
  }

  public async getMyDay(
    now: Date = new Date(),
    queryOptions?: IMyDayQueryOptions
  ): Promise<IMyDayData> {
    const location = queryOptions?.location;
    const key = `${location?.city || ''}|${location?.country || ''}`;

    if (this.cache && this.cache.key === key && Date.now() - this.cache.at < CACHE_TTL_MS) {
      return this.cache.data;
    }

    const data = this.load(now, queryOptions);
    this.cache = { key, at: Date.now(), data };
    // A failed load must not be replayed for the full TTL — let the next call
    // (retry, or the inline ↔ fullscreen switch) try again from scratch.
    data.catch(() => {
      if (this.cache && this.cache.data === data) {
        this.cache = undefined;
      }
    });
    return data;
  }

  private getClient(): Promise<MSGraphClientV3> {
    return withTimeout(
      this.clientFactory.getClient('3'),
      CLIENT_TIMEOUT_MS,
      'Timed out getting a Microsoft Graph client (sign-in did not complete in time).'
    );
  }

  private async load(now: Date, queryOptions?: IMyDayQueryOptions): Promise<IMyDayData> {
    const issues: IMyDayDataIssue[] = [];

    // TEMP: only loading calendar items right now to isolate load-time issues.
    // Tasks/mail/news/weather are skipped below (loadGraphSections / weather
    // section) until this is reverted.
    const [{ meetings, tasks, mail, news }, weather] = await Promise.all([
      this.loadGraphSections(now, issues),
      Promise.resolve(undefined)
    ]);

    return {
      user: this.options.currentUser,
      meetings,
      tasks,
      mail,
      news,
      weather,
      quickActions: mockQuickActions,
      issues: issues.length > 0 ? issues : undefined
    };
  }

  /**
   * Loads the four sections that need a signed-in Graph client. If the client
   * itself cannot be acquired (e.g. a timed-out/failed silent sign-in), each
   * section is recorded as its own issue and resolves empty instead of
   * rejecting the whole load — the dashboard still renders with weather intact.
   */
  private async loadGraphSections(
    now: Date,
    issues: IMyDayDataIssue[]
  ): Promise<Pick<IMyDayData, 'meetings' | 'tasks' | 'mail' | 'news'>> {
    let client: MSGraphClientV3;
    try {
      console.log('[My Day] Acquiring a Microsoft Graph client…');
      client = await this.getClient();
    } catch (error) {
      console.warn('[My Day] Failed to acquire a Microsoft Graph client:', error);
      (['meetings', 'tasks', 'mail', 'news'] as const).forEach((section) => {
        issues.push({
          section,
          message: `${SECTION_LABELS[section]} could not be loaded right now.`
        });
      });
      return { meetings: [], tasks: [], mail: [], news: [] };
    }

    // TEMP: only loading calendar items right now to isolate load-time issues.
    const meetings = await this.section('meetings', issues, [], async () =>
      (await fetchTodaysEvents(client, now))
        .map(mapMeeting)
        .sort((a, b) => a.start.getTime() - b.start.getTime())
    );

    return { meetings, tasks: [], mail: [], news: [] };
  }

  /** Runs one section, recording an issue and returning `fallback` on failure. */
  private async section<T>(
    section: MyDaySection,
    issues: IMyDayDataIssue[],
    fallback: T,
    run: () => Promise<T>
  ): Promise<T> {
    console.log(`[My Day] Loading ${section}…`);
    try {
      return await run();
    } catch (error) {
      console.warn(`[My Day] Failed to load ${section}:`, error);
      issues.push({
        section,
        message: `${SECTION_LABELS[section]} could not be loaded right now.`
      });
      return fallback;
    }
  }
}
