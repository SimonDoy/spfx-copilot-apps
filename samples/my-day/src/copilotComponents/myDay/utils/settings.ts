import * as React from 'react';

import type { MyDayDataSource } from '../services/IMyDayDataService';

/** Identifiers for the dashboard panels the user can show or hide. */
export const PANEL_IDS = [
  'agenda',
  'tasks',
  'mail',
  'news',
  'quickActions',
  'planMyDay'
] as const;

export type PanelId = (typeof PANEL_IDS)[number];

/**
 * User-adjustable settings for the My Day dashboard.
 *
 * In this demo these are **intentionally** persisted to `sessionStorage` so the
 * values survive re-renders and reloads within the current browser session but
 * are deliberately forgotten when the session ends — no server-side storage is
 * involved.
 */
export interface IMyDaySettings {
  /** When `true`, temperatures are shown primarily in Fahrenheit. */
  useFahrenheit: boolean;
  /** Preferred city. Drives the live weather lookup in `graph` mode. */
  city: string;
  /** Preferred country, used to disambiguate the city when geocoding. */
  country: string;
  /** Panels currently shown in the full-screen dashboard. */
  visiblePanels: PanelId[];
  /**
   * Session override for the data source. `undefined` means "not chosen by the
   * user", so the component's `dataSource` property wins.
   */
  dataSource?: MyDayDataSource;
}

export const DEFAULT_SETTINGS: IMyDaySettings = {
  useFahrenheit: false,
  city: 'Redmond',
  country: 'United States',
  visiblePanels: [...PANEL_IDS]
};

const STORAGE_KEY = 'myDay.settings';

const getSessionStorage = (): Storage | undefined => {
  try {
    return typeof window !== 'undefined' ? window.sessionStorage : undefined;
  } catch {
    // Access to sessionStorage can throw in sandboxed/embedded hosts.
    return undefined;
  }
};

/** Keeps only known panel ids, preserving canonical order. */
const sanitizePanels = (panels: unknown): PanelId[] => {
  if (!Array.isArray(panels)) {
    return [...DEFAULT_SETTINGS.visiblePanels];
  }
  return PANEL_IDS.filter((id) => panels.indexOf(id) !== -1);
};

const sanitizeDataSource = (value: unknown): MyDayDataSource | undefined =>
  value === 'graph' || value === 'mock' ? value : undefined;

/**
 * The data source actually in effect: the user's session choice when they have
 * made one, otherwise the value configured on the component.
 */
export const resolveDataSource = (
  settings: IMyDaySettings,
  configured?: MyDayDataSource
): MyDayDataSource => settings.dataSource || configured || 'mock';

/** Reads the settings from session storage, falling back to defaults. */
export const loadSettings = (): IMyDaySettings => {
  const storage = getSessionStorage();
  if (!storage) {
    return { ...DEFAULT_SETTINGS };
  }

  try {
    const raw = storage.getItem(STORAGE_KEY);
    if (!raw) {
      return { ...DEFAULT_SETTINGS };
    }
    const parsed = JSON.parse(raw) as Partial<IMyDaySettings>;
    return {
      ...DEFAULT_SETTINGS,
      ...parsed,
      visiblePanels: sanitizePanels(parsed.visiblePanels),
      dataSource: sanitizeDataSource(parsed.dataSource)
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
};

/** Persists the settings to session storage (best-effort). */
export const saveSettings = (settings: IMyDaySettings): void => {
  const storage = getSessionStorage();
  if (!storage) {
    return;
  }

  try {
    storage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // Ignore quota / access errors — settings are non-critical in the demo.
  }
};

/**
 * Session-backed settings hook. Initializes from `sessionStorage`, and writes
 * every change back so it is shared across the fullscreen view for the current
 * session only.
 */
export const useMyDaySettings = (): [
  IMyDaySettings,
  (patch: Partial<IMyDaySettings>) => void
] => {
  const [settings, setSettings] = React.useState<IMyDaySettings>(loadSettings);

  const updateSettings = React.useCallback((patch: Partial<IMyDaySettings>): void => {
    setSettings((prev) => {
      const next = { ...prev, ...patch };
      saveSettings(next);
      return next;
    });
  }, []);

  return [settings, updateSettings];
};
