import * as React from 'react';

import type { SPCopilotDisplayMode, SPCopilotTheme } from '@microsoft/sp-copilot-component';

import type { IMyDayCopilotComponentProperties } from '../MyDayCopilotComponentProperties';
import type { IUser } from '../models/myDay';
import type { IMyDayDataService, MyDayDataSource } from '../services/IMyDayDataService';
import { MockMyDayDataService } from '../services/MockMyDayDataService';
import MyDayErrorBoundary from './MyDayErrorBoundary';
import MyDayFullscreen from './MyDayFullscreen';
import MyDayInline from './MyDayInline';
import MyDayThemeProvider from './MyDayThemeProvider';

/**
 * Props for the {@link MyDayApp} root React component.
 *
 * Tool inputs supplied by Copilot ({@link IMyDayCopilotComponentProperties}) are
 * spread in directly, while host state (theme, display mode) and host-driven
 * actions are passed alongside them. The component derives all of its UI from
 * these props — it never mirrors host state internally.
 */
export interface IMyDayAppProps extends IMyDayCopilotComponentProperties {
  /** Signed-in user resolved from the SPFx page context. */
  currentUser?: IUser;
  /**
   * Builds (and caches) the data service for a source. Supplied by the SPFx
   * component so the React tree never touches `@microsoft/sp-http`.
   */
  createDataService?: (source: MyDayDataSource) => IMyDayDataService;
  /** Color theme advertised by the Copilot host. */
  theme?: SPCopilotTheme;
  /** Container display mode advertised by the Copilot host. */
  displayMode?: SPCopilotDisplayMode;
  /** Display modes the host supports. */
  availableDisplayModes?: SPCopilotDisplayMode[];
  /** Requests the host to switch the component into fullscreen. */
  onRequestFullscreen?: () => void;
}

/** Used when no factory is supplied (e.g. tests, storybook-style hosting). */
const mockOnly = (): IMyDayDataService => new MockMyDayDataService();

/**
 * Baseline root component rendered into the Copilot host DOM element.
 *
 * Selects the display-mode-specific view ({@link MyDayInline} or
 * {@link MyDayFullscreen}) from the host-advertised display mode.
 */
const MyDayApp: React.FunctionComponent<IMyDayAppProps> = (props) => {
  const {
    message,
    dataSource,
    currentUser,
    createDataService = mockOnly,
    theme,
    displayMode,
    onRequestFullscreen
  } = props;

  const isFullscreen = displayMode === 'fullscreen';

  const view = isFullscreen ? (
    <MyDayFullscreen
      message={message}
      dataSource={dataSource}
      currentUser={currentUser}
      createDataService={createDataService}
      theme={theme}
    />
  ) : (
    <MyDayInline
      message={message}
      dataSource={dataSource}
      currentUser={currentUser}
      createDataService={createDataService}
      theme={theme}
      onRequestFullscreen={onRequestFullscreen}
    />
  );

  return (
    <MyDayThemeProvider theme={theme}>
      <MyDayErrorBoundary compact={!isFullscreen}>{view}</MyDayErrorBoundary>
    </MyDayThemeProvider>
  );
};

export default MyDayApp;
