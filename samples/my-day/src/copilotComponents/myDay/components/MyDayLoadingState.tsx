import * as React from 'react';

import { makeStyles, Spinner, tokens, Text } from '@fluentui/react-components';

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
    padding: '48px 24px',
    textAlign: 'center',
    fontFamily: tokens.fontFamilyBase
  },
  compact: {
    padding: '32px 0'
  },
  label: {
    color: tokens.colorNeutralForeground2
  }
});

export interface IMyDayLoadingStateProps {
  /** Message shown under the spinner. Defaults to a friendly, branded line. */
  label?: string;
  /** Renders the tighter variant used by the inline card. */
  compact?: boolean;
}

/**
 * Shared "component is loading" surface for the inline and full-screen views,
 * shown while the initial `useMyDayData` fetch is in flight.
 */
const MyDayLoadingState: React.FunctionComponent<IMyDayLoadingStateProps> = (props) => {
  const styles = useStyles();
  const { label = 'Putting your day together…', compact } = props;

  return (
    <div className={compact ? `${styles.root} ${styles.compact}` : styles.root} role="status">
      <Spinner size={compact ? 'small' : 'medium'} />
      <Text className={styles.label} size={compact ? 200 : 300}>
        {label}
      </Text>
    </div>
  );
};

export default MyDayLoadingState;
