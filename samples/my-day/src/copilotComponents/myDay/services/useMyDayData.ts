import * as React from 'react';

import type { IMyDayData, IUser } from '../models/myDay';
import type { IMyDayDataService, IMyDayQueryOptions } from './IMyDayDataService';
import { buildMockMyDay } from './MockMyDayDataService';

export interface IMyDayDataState {
  data: IMyDayData;
  loading: boolean;
}

/**
 * Loads the My Day aggregate for the given service and keeps the view in a
 * simple `{ data, loading }` state.
 *
 * The mock provider resolves in a microtask, so the demo shows no perceptible
 * loading state. A rejected load (only possible if the whole Graph client fails
 * to initialise) falls back to the mock data with an explanatory issue, so the
 * dashboard is never blank.
 */
export const useMyDayData = (
  service: IMyDayDataService,
  now: Date,
  currentUser?: IUser,
  options?: IMyDayQueryOptions
): IMyDayDataState => {
  const city = options?.location?.city;
  const country = options?.location?.country;

  const [state, setState] = React.useState<IMyDayDataState>(() => ({
    data: buildMockMyDay(now),
    loading: true
  }));

  React.useEffect(() => {
    let cancelled = false;
    setState((prev) => ({ ...prev, loading: true }));

    const location = city ? { city, country: country || '' } : undefined;

    service
      .getMyDay(now, { location })
      .catch((error: unknown) => {
        // eslint-disable-next-line no-console
        console.warn('[My Day] Falling back to demo data:', error);
        return {
          ...buildMockMyDay(now),
          issues: [
            {
              section: 'meetings' as const,
              message: 'Live data is unavailable — showing demo data.'
            }
          ]
        };
      })
      .then((data) => {
        if (!cancelled) {
          setState({ data, loading: false });
        }
      })
      .catch(() => {
        /* `catch` above already resolved; nothing can reach here. */
      });

    return () => {
      cancelled = true;
    };
  }, [service, now, city, country]);

  // Prefer the real signed-in user from the page context over the mock user,
  // keeping the mock user as a fallback (e.g. tests without a page context).
  const data = React.useMemo(
    () => (currentUser ? { ...state.data, user: currentUser } : state.data),
    [state.data, currentUser]
  );

  return { data, loading: state.loading };
};
