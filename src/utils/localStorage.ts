import { Dataset, ReportingWidget, WidgetWidth } from '../types/dashboard';

const STORAGE_KEYS = {
  DATASETS: 'xlsm_studio_datasets_v3',
  WIDGETS: 'xlsm_studio_widgets_v3',
  LAYOUT_CONFIG: 'xlsm_studio_layout_config_v1'
};

export interface DashboardLayoutConfig {
  order: string[]; // Ordered list of widget IDs
  widths: Record<string, WidgetWidth>; // Widget ID -> WidgetWidth
  lastModified: string;
}

export function saveLocalDatasets(datasets: Dataset[]) {
  try {
    const serializable = datasets.map((d) => ({
      ...d,
      rawZip: undefined,
      unprotectedBlob: undefined
    }));
    localStorage.setItem(STORAGE_KEYS.DATASETS, JSON.stringify(serializable));
  } catch (err) {
    console.warn('LocalStorage write error:', err);
  }
}

export function getLocalDatasets(): Dataset[] | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.DATASETS);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export function saveLocalWidgets(widgets: ReportingWidget[]) {
  try {
    localStorage.setItem(STORAGE_KEYS.WIDGETS, JSON.stringify(widgets));
    // Also update layout config snapshot
    const layoutConfig: DashboardLayoutConfig = {
      order: widgets.map((w) => w.id),
      widths: widgets.reduce((acc, w) => {
        acc[w.id] = w.width;
        return acc;
      }, {} as Record<string, WidgetWidth>),
      lastModified: new Date().toISOString()
    };
    localStorage.setItem(STORAGE_KEYS.LAYOUT_CONFIG, JSON.stringify(layoutConfig));
  } catch (err) {
    console.warn('LocalStorage write error for widgets:', err);
  }
}

export function getLocalWidgets(): ReportingWidget[] | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.WIDGETS);
    if (!raw) return null;
    const parsed: ReportingWidget[] = JSON.parse(raw);
    return sortAndApplyLayout(parsed);
  } catch {
    return null;
  }
}

export function getLayoutConfig(): DashboardLayoutConfig | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.LAYOUT_CONFIG);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export function saveLayoutConfig(config: DashboardLayoutConfig) {
  try {
    localStorage.setItem(STORAGE_KEYS.LAYOUT_CONFIG, JSON.stringify(config));
  } catch (err) {
    console.warn('LocalStorage write error for layout config:', err);
  }
}

export function resetWidgetLayoutStorage() {
  try {
    localStorage.removeItem(STORAGE_KEYS.WIDGETS);
    localStorage.removeItem(STORAGE_KEYS.LAYOUT_CONFIG);
  } catch (err) {
    console.warn('LocalStorage error resetting widget layout:', err);
  }
}

/**
 * Re-orders and applies saved widget widths according to stored layout configuration.
 */
export function sortAndApplyLayout(widgets: ReportingWidget[]): ReportingWidget[] {
  const layout = getLayoutConfig();
  if (!layout || !layout.order || layout.order.length === 0) {
    return widgets;
  }

  const widgetMap = new Map<string, ReportingWidget>();
  widgets.forEach((w) => {
    const updatedWidth = layout.widths?.[w.id] || w.width;
    widgetMap.set(w.id, { ...w, width: updatedWidth });
  });

  const ordered: ReportingWidget[] = [];
  layout.order.forEach((id) => {
    const found = widgetMap.get(id);
    if (found) {
      ordered.push(found);
      widgetMap.delete(id);
    }
  });

  // Append any newly added widgets not present in the order history
  widgetMap.forEach((remaining) => {
    ordered.push(remaining);
  });

  return ordered;
}

