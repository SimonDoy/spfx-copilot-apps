import * as React from 'react';

import MyDayErrorState from './MyDayErrorState';

export interface IMyDayErrorBoundaryProps {
  /** Renders the tighter error variant used by the inline card. */
  compact?: boolean;
  children: React.ReactNode;
}

interface IMyDayErrorBoundaryState {
  error?: Error;
  /** Bumped on retry so the subtree remounts from scratch. */
  attempt: number;
}

/**
 * Catches render-time failures anywhere in the My Day tree and replaces them
 * with {@link MyDayErrorState}, so a bad payload or a broken child can never
 * leave the Copilot host with a blank surface. Retry remounts the subtree.
 */
export default class MyDayErrorBoundary extends React.Component<
  IMyDayErrorBoundaryProps,
  IMyDayErrorBoundaryState
> {
  public constructor(props: IMyDayErrorBoundaryProps) {
    super(props);
    this.state = { attempt: 0 };
  }

  public static getDerivedStateFromError(error: Error): Partial<IMyDayErrorBoundaryState> {
    return { error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('[My Day] Render failed:', error, info.componentStack);
  }

  public render(): React.ReactNode {
    const { compact, children } = this.props;
    const { error, attempt } = this.state;

    if (error) {
      return (
        <MyDayErrorState
          compact={compact}
          details={error.stack || error.message}
          onRetry={this._retry}
        />
      );
    }

    return <React.Fragment key={attempt}>{children}</React.Fragment>;
  }

  private _retry = (): void => {
    this.setState((prev) => ({ error: undefined, attempt: prev.attempt + 1 }));
  };
}
