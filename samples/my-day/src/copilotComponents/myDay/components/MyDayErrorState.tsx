import * as React from 'react';

import { Button, makeStyles, tokens, Text } from '@fluentui/react-components';
import { ArrowClockwise20Regular, WeatherMoon24Regular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '12px',
    width: '100%',
    boxSizing: 'border-box',
    minWidth: 0,
    padding: '32px 24px',
    textAlign: 'center',
    fontFamily: tokens.fontFamilyBase,
    borderRadius: tokens.borderRadiusXLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow2
  },
  compact: {
    padding: '20px 16px',
    gap: '8px',
    boxShadow: 'none'
  },
  badge: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '48px',
    height: '48px',
    borderRadius: tokens.borderRadiusCircular,
    color: tokens.colorNeutralForeground2,
    backgroundColor: tokens.colorNeutralBackground3
  },
  title: {
    fontWeight: tokens.fontWeightSemibold
  },
  body: {
    maxWidth: '440px',
    color: tokens.colorNeutralForeground2
  },
  details: {
    maxWidth: '440px',
    width: '100%',
    marginTop: '4px',
    textAlign: 'left',
    color: tokens.colorNeutralForeground3
  },
  summary: {
    cursor: 'pointer',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3
  },
  pre: {
    margin: '8px 0 0',
    padding: '8px',
    maxHeight: '140px',
    overflow: 'auto',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3
  }
});

export interface IMyDayErrorStateProps {
  /** Short headline. Defaults to a friendly, non-technical message. */
  title?: string;
  /** Explanatory line under the headline. */
  description?: string;
  /** Technical detail, shown collapsed so the surface stays calm. */
  details?: string;
  /** Renders the tighter variant used by the inline card. */
  compact?: boolean;
  onRetry?: () => void;
}

/**
 * Friendly, theme-aware failure surface. Deliberately reuses the dashboard card
 * chrome so a failure still looks like part of My Day rather than a crash.
 */
const MyDayErrorState: React.FunctionComponent<IMyDayErrorStateProps> = (props) => {
  const styles = useStyles();
  const {
    title = 'My Day is taking a break',
    description = 'Something went wrong while putting your day together. Nothing has been changed — try again in a moment.',
    details,
    compact,
    onRetry
  } = props;

  return (
    <div
      className={compact ? `${styles.root} ${styles.compact}` : styles.root}
      role="alert"
      data-my-day-error="true"
    >
      <span className={styles.badge} aria-hidden="true">
        <WeatherMoon24Regular />
      </span>

      <Text className={styles.title} size={compact ? 400 : 500}>
        {title}
      </Text>
      <Text className={styles.body} size={compact ? 200 : 300}>
        {description}
      </Text>

      {onRetry && (
        <Button appearance="primary" icon={<ArrowClockwise20Regular />} onClick={onRetry}>
          Try again
        </Button>
      )}

      {details && (
        <details className={styles.details}>
          <summary className={styles.summary}>Technical details</summary>
          <pre className={styles.pre}>{details}</pre>
        </details>
      )}
    </div>
  );
};

export default MyDayErrorState;
