import * as React from 'react';

import { makeStyles, Spinner, tokens } from '@fluentui/react-components';
import type { SPCopilotTheme } from '@microsoft/sp-copilot-component';

import type { IMyDayCopilotComponentProperties } from '../MyDayCopilotComponentProperties';
import type { IUser } from '../models/myDay';
import type { IMyDayDataService, MyDayDataSource } from '../services/IMyDayDataService';
import { useMyDayData } from '../services/useMyDayData';
import { fadeIn } from '../utils/motion';
import { resolveDataSource, useMyDaySettings } from '../utils/settings';
import AgendaTimeline from './fullscreen/AgendaTimeline';
import TasksPanel from './fullscreen/TasksPanel';
import InlineDetailHeader from './inline/InlineDetailHeader';
import InlineSummary from './inline/InlineSummary';
import NewsList from './inline/NewsList';
import type { InlineView } from './inline/views';

/**
 * Props for the {@link MyDayInline} view rendered when the host display mode is
 * `'inline'`.
 */
export interface IMyDayInlineProps extends IMyDayCopilotComponentProperties {
  /** Signed-in user resolved from the SPFx page context. */
  currentUser?: IUser;
  /** Builds (and caches) the data service for a source. */
  createDataService: (source: MyDayDataSource) => IMyDayDataService;
  /** Color theme advertised by the Copilot host. */
  theme?: SPCopilotTheme;
  /** Requests the host to switch the component into fullscreen. */
  onRequestFullscreen?: () => void;
}

const useStyles = makeStyles({
  surface: {
    width: '100%',
    boxSizing: 'border-box',
    minWidth: 0,
    padding: '12px',
    fontFamily: tokens.fontFamilyBase
  },
  enter: {
    animationName: fadeIn,
    animationDuration: tokens.durationSlow,
    animationTimingFunction: tokens.curveDecelerateMid,
    animationFillMode: 'both',
    '@media (prefers-reduced-motion: reduce)': {
      animationName: 'none',
      animationDuration: '1ms'
    }
  },
  loading: {
    display: 'flex',
    justifyContent: 'center',
    padding: '32px 0'
  }
});

/**
 * Inline display-mode view. Owns the local drill-down navigation state and
 * renders the time-aware summary plus meetings / tasks / news detail views. The
 * meetings and tasks drill-downs reuse the full-screen `AgendaTimeline` and
 * `TasksPanel` controls (news keeps its own inline listing). Theme is derived
 * from the host-advertised `theme` prop (never mirrored).
 */
const MyDayInline: React.FunctionComponent<IMyDayInlineProps> = (props) => {
  const styles = useStyles();
  const { currentUser, createDataService, dataSource, onRequestFullscreen } = props;

  const [view, setView] = React.useState<InlineView>('summary');
  const [settings] = useMyDaySettings();

  // Resolve the data once per mount against the current clock so relative times
  // stay live and future-biased.
  const now = React.useMemo(() => new Date(), []);
  const source = resolveDataSource(settings, dataSource);
  const service = React.useMemo(
    () => createDataService(source),
    [createDataService, source]
  );
  const location = React.useMemo(
    () => ({ city: settings.city, country: settings.country }),
    [settings.city, settings.country]
  );
  const { data, loading } = useMyDayData(service, now, currentUser, { location });

  const back = React.useCallback(() => setView('summary'), []);

  let content: React.ReactNode;
  if (loading) {
    content = (
      <div className={styles.loading}>
        <Spinner size="tiny" label="Loading your day…" />
      </div>
    );
  } else {
    switch (view) {
      case 'meetings':
        content = (
          <>
            <InlineDetailHeader onBack={back} />
            <AgendaTimeline meetings={data.meetings} now={now} />
          </>
        );
        break;
      case 'tasks':
        content = (
          <>
            <InlineDetailHeader onBack={back} />
            <TasksPanel tasks={data.tasks} now={now} />
          </>
        );
        break;
      case 'news':
        content = <NewsList news={data.news} now={now} onBack={back} />;
        break;
      default:
        content = (
          <InlineSummary
            data={data}
            now={now}
            onNavigate={setView}
            onRequestFullscreen={onRequestFullscreen}
          />
        );
    }
  }

  return (
    <div className={styles.surface} data-display-mode="inline">
      {view === 'summary' || loading ? (
        content
      ) : (
        <div key={view} className={styles.enter}>
          {content}
        </div>
      )}
    </div>
  );
};

export default MyDayInline;
