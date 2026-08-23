// Factory that wires a My Day data service to the Copilot host.
//
// This is the ONLY file that touches @microsoft/sp-http, so both services stay
// free of SPFx types and trivially unit-testable with plain stubs.

import type { CopilotComponentContext } from '@microsoft/sp-copilot-component';
import { HttpClient } from '@microsoft/sp-http';

import type { IUser } from '../models/myDay';
import { GraphMyDayDataService } from './GraphMyDayDataService';
import type { IMyDayDataService, MyDayDataSource } from './IMyDayDataService';
import { MockMyDayDataService } from './MockMyDayDataService';
import { OpenMeteoWeatherService, type JsonGet } from './OpenMeteoWeatherService';

export interface ICreateMyDayDataServiceOptions {
  /** Signed-in user resolved from the SPFx page context. */
  currentUser: IUser;
  /** SharePoint site to read news from. Defaults to the tenant root site. */
  newsSiteUrl?: string;
}

/**
 * Builds the data service for the requested source. `mock` needs no context and
 * no permissions; `graph` reads live data as the signed-in user.
 */
export const createMyDayDataService = (
  source: MyDayDataSource,
  context: CopilotComponentContext,
  options: ICreateMyDayDataServiceOptions
): IMyDayDataService => {
  if (source !== 'graph') {
    return new MockMyDayDataService();
  }

  // Anonymous GET against the public weather API — no SharePoint or Graph
  // credentials are attached to these requests.
  const getJson: JsonGet = async (url) => {
    const response = await context.httpClient.get(url, HttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Weather request failed with ${response.status}`);
    }

    // A body that is empty or not valid JSON (e.g. a proxy/CORS hiccup
    // returning no content) must surface as a clean error, not a raw
    // "Unexpected end of JSON..." SyntaxError from `.json()`.
    const text = await response.text();
    if (!text) {
      throw new Error('Weather request returned an empty response.');
    }
    try {
      return JSON.parse(text);
    } catch {
      throw new Error('Weather request returned an invalid response.');
    }
  };

  return new GraphMyDayDataService(context.msGraphClientFactory, {
    currentUser: options.currentUser,
    webAbsoluteUrl: context.pageContext.web.absoluteUrl,
    newsSiteUrl: options.newsSiteUrl,
    weatherService: new OpenMeteoWeatherService(getJson)
  });
};
