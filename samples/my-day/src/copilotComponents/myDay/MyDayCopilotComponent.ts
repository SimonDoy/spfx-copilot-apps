import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { BaseCopilotComponent } from '@microsoft/sp-copilot-component';

import MyDayApp from './components/MyDayApp';
import type { IMyDayCopilotComponentProperties } from './MyDayCopilotComponentProperties';
import { resolveCurrentUser } from './services/CurrentUserService';
import { createMyDayDataService } from './services/createMyDayDataService';
import type { IMyDayDataService, MyDayDataSource } from './services/IMyDayDataService';

export default class MyDayCopilotComponent extends BaseCopilotComponent<IMyDayCopilotComponentProperties> {
  private readonly _services: Map<MyDayDataSource, IMyDayDataService> = new Map();

  protected render(): void {
    const currentUser = resolveCurrentUser(this.context);

    const element: React.ReactElement = React.createElement(MyDayApp, {
      message: this.properties.message,
      dataSource: this.properties.dataSource,
      newsSiteUrl: this.properties.newsSiteUrl,
      currentUser,
      createDataService: (source: MyDayDataSource) => this._resolveService(source, currentUser),
      theme: this.hostContext.theme,
      displayMode: this.hostContext.displayMode,
      availableDisplayModes: this.hostContext.availableDisplayModes,
      onRequestFullscreen: this._handleRequestFullscreen
    });

    ReactDOM.render(element, this.context.domElement);

  }

  protected onTeardown(reason: string | undefined): Promise<void> {
    ReactDOM.unmountComponentAtNode(this.context.domElement);
    return super.onTeardown(reason);
  }

  /** Services are cached per source so switching back and forth reuses caches. */
  private _resolveService(
    source: MyDayDataSource,
    currentUser: ReturnType<typeof resolveCurrentUser>
  ): IMyDayDataService {
    const existing = this._services.get(source);
    if (existing) {
      return existing;
    }

    const service = createMyDayDataService(source, this.context, {
      currentUser,
      newsSiteUrl: this.properties.newsSiteUrl
    });
    this._services.set(source, service);
    return service;
  }

  private _handleRequestFullscreen = (): void => {
    this.requestDisplayModeAsync('fullscreen').catch(() => {
      /* Host rejected or is unavailable; host context stays unchanged. */
    });
  };
}
