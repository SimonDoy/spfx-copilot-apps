import type { IMyDayData } from '../models/myDay';

/** Where My Day reads its data from. */
export type MyDayDataSource = 'mock' | 'graph';

/** Optional per-call inputs the live providers use to shape their queries. */
export interface IMyDayQueryOptions {
  /** User's preferred location, used by the weather provider. */
  location?: { city: string; country: string };
}

/**
 * Contract for supplying My Day data. Two implementations ship: a mock provider
 * and a Microsoft Graph provider. Both return the same `IMyDayData` view models,
 * so the UI never knows which one it is talking to.
 */
export interface IMyDayDataService {
  /**
   * Returns the aggregated My Day data. `now` lets callers resolve relative
   * times deterministically (and keeps the demo live at render time).
   */
  getMyDay(now?: Date, options?: IMyDayQueryOptions): Promise<IMyDayData>;
}
