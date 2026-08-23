import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { BaseCopilotComponent } from '@microsoft/sp-copilot-component';

import MyDayApp from './components/MyDayApp';
import MyDayErrorState from './components/MyDayErrorState';
import MyDayThemeProvider from './components/MyDayThemeProvider';
import type { IMyDayCopilotComponentProperties } from './MyDayCopilotComponentProperties';
import { resolveCurrentUser } from './services/CurrentUserService';
import { createMyDayDataService } from './services/createMyDayDataService';
import type { IMyDayDataService, MyDayDataSource } from './services/IMyDayDataService';
import { MockMyDayDataService } from './services/MockMyDayDataService';

export default class MyDayCopilotComponent extends BaseCopilotComponent<IMyDayCopilotComponentProperties> {
  private readonly _services: Map<MyDayDataSource, IMyDayDataService> = new Map();

  public constructor() {
    super();
    console.log('[My Day] MyDayCopilotComponent created');
  }

  protected render(): void {
    let element: React.ReactElement;

    // The React tree has its own error boundary; this guards the setup done
    // here (page context, service factory) so the host never sees a throw.
    try {
      element = this._buildApp();
    } catch (error) {
      console.error('[My Day] Failed to build the component:', error);
      element = this._buildFailure(error);
    }

    ReactDOM.render(element, this.context.domElement);
  }

  protected onTeardown(reason: string | undefined): Promise<void> {
    ReactDOM.unmountComponentAtNode(this.context.domElement);
    return super.onTeardown(reason);
  }

  private _buildApp(): React.ReactElement {
    const currentUser = resolveCurrentUser(this.context);

    return React.createElement(MyDayApp, {
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
  }

  /** Last-resort surface when the component cannot even be constructed. */
  private _buildFailure(error: unknown): React.ReactElement {
    const details = error instanceof Error ? error.stack || error.message : String(error);

    return React.createElement(
      MyDayThemeProvider,
      { theme: this.hostContext.theme },
      React.createElement(MyDayErrorState, {
        compact: this.hostContext.displayMode !== 'fullscreen',
        details,
        onRetry: () => this.render()
      })
    );
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

    let service: IMyDayDataService;
    try {
      service = createMyDayDataService(source, this.context, {
        currentUser,
        newsSiteUrl: this.properties.newsSiteUrl
      });
    } catch (error) {
      // A missing Graph client or malformed context must not break the render;
      // demo data keeps the dashboard usable.
      console.warn('[My Day] Falling back to demo data service:', error);
      service = new MockMyDayDataService();
    }

    this._services.set(source, service);
    return service;
  }

  private _handleRequestFullscreen = (): void => {
    this.requestDisplayModeAsync('fullscreen').catch(() => {
      /* Host rejected or is unavailable; host context stays unchanged. */
    });
  };
}
