import type { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

import { mockQuickActions } from '../mockData';
import type { IMyDayData, IMyDayDataIssue, IUser, MyDaySection } from '../models/myDay';
import { buildUserPhotoUrl } from './CurrentUserService';
import {
  fetchImportantMail,
  fetchNews,
  fetchOpenTasks,
  fetchTodaysEvents,
  resolveSiteId
} from './graphQueries';
import type { IMyDayDataService, IMyDayQueryOptions } from './IMyDayDataService';
import { mapMail, mapMeeting, mapNews, mapTask } from './mappers';
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
  private clientPromise: Promise<MSGraphClientV3> | undefined;
  private cache: { key: string; at: number; data: Promise<IMyDayData> } | undefined;

  public constructor(
    clientFactory: MSGraphClientFactory,
    options: IGraphMyDayDataServiceOptions
  ) {
    this.clientFactory = clientFactory;
    this.options = options;
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
    return data;
  }

  private getClient(): Promise<MSGraphClientV3> {
    const existing = this.clientPromise;
    if (existing) {
      return existing;
    }

    const created = this.clientFactory.getClient('3');
    this.clientPromise = created;
    return created;
  }

  private async load(now: Date, queryOptions?: IMyDayQueryOptions): Promise<IMyDayData> {
    const issues: IMyDayDataIssue[] = [];
    const client = await this.getClient();

    const photoUrl = (email: string): string | undefined =>
      buildUserPhotoUrl(this.options.webAbsoluteUrl, email, 'S');

    const [meetings, tasks, mail, news, weather] = await Promise.all([
      this.section('meetings', issues, [], async () =>
        (await fetchTodaysEvents(client, now))
          .map(mapMeeting)
          .sort((a, b) => a.start.getTime() - b.start.getTime())
      ),
      this.section('tasks', issues, [], async () =>
        (await fetchOpenTasks(client)).map(mapTask)
      ),
      this.section('mail', issues, [], async () =>
        (await fetchImportantMail(client))
          .map((m) => ({ ...mapMail(m), senderPhotoUrl: photoUrl(m.from.emailAddress.address) }))
          .sort((a, b) => b.receivedAt.getTime() - a.receivedAt.getTime())
      ),
      this.section('news', issues, [], async () => {
        const siteId = await resolveSiteId(client, this.options.newsSiteUrl);
        if (!siteId) {
          return [];
        }
        return (await fetchNews(client, siteId, photoUrl))
          .map(mapNews)
          .sort((a, b) => b.publishedAt.getTime() - a.publishedAt.getTime());
      }),
      this.section('weather', issues, undefined, async () => {
        const location = queryOptions?.location;
        if (!this.options.weatherService || !location?.city) {
          return undefined;
        }
        return this.options.weatherService.getWeather(location.city, location.country);
      })
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

  /** Runs one section, recording an issue and returning `fallback` on failure. */
  private async section<T>(
    section: MyDaySection,
    issues: IMyDayDataIssue[],
    fallback: T,
    run: () => Promise<T>
  ): Promise<T> {
    try {
      return await run();
    } catch (error) {
      // eslint-disable-next-line no-console
      console.warn(`[My Day] Failed to load ${section}:`, error);
      issues.push({
        section,
        message: `${SECTION_LABELS[section]} could not be loaded right now.`
      });
      return fallback;
    }
  }
}
