import { mockEvents, mockMail, mockNews, mockQuickActions, mockTasks, mockUser, mockWeather } from '../mockData';
import type { IMyDayData } from '../models/myDay';
import type { IMyDayDataService } from './IMyDayDataService';
import { mapMail, mapMeeting, mapNews, mapTask, mapUser } from './mappers';
import { resolveEvents, resolveMail, resolveNews, resolveTasks } from './resolveMockData';

/**
 * Builds the mock aggregate synchronously: resolves the bundled mock data's
 * relative time offsets against `now`, then maps the Graph-shaped objects to
 * view models — so the demo is always live and future-biased.
 *
 * Also used as the render-ready starting state and as the fallback when live
 * data cannot be loaded at all.
 */
export const buildMockMyDay = (now: Date = new Date()): IMyDayData => {
  const meetings = resolveEvents(mockEvents, now)
    .map(mapMeeting)
    .sort((a, b) => a.start.getTime() - b.start.getTime());

  const tasks = resolveTasks(mockTasks, now).map(mapTask);

  const news = resolveNews(mockNews, now)
    .map(mapNews)
    .sort((a, b) => b.publishedAt.getTime() - a.publishedAt.getTime());

  const mail = resolveMail(mockMail, now)
    .map(mapMail)
    .sort((a, b) => b.receivedAt.getTime() - a.receivedAt.getTime());

  return {
    user: mapUser(mockUser),
    meetings,
    tasks,
    news,
    mail,
    weather: mockWeather,
    quickActions: mockQuickActions
  };
};

/** Mock implementation of {@link IMyDayDataService}. Needs no permissions. */
export class MockMyDayDataService implements IMyDayDataService {
  public async getMyDay(now: Date = new Date()): Promise<IMyDayData> {
    return buildMockMyDay(now);
  }
}
