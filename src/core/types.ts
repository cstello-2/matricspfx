export type Row = Record<string, unknown>;

export interface MetricsConfig {
  $schema?: string;
  title?: string;
  description?: string;
  refreshSeconds?: number;
  dataSources: Record<string, DataSourceConfig>;
  datasets?: Record<string, DatasetConfig>;
  widgets: WidgetConfig[];
  list?: ListViewConfig;
  dataView?: false | DataViewConfig;
  layout?: LayoutConfig;
  theme?: ThemeConfig;
  locale?: string;
}

export interface LayoutConfig {
  columns?: number;
  gap?: number;
}

export interface ThemeConfig {
  palette?: 'default' | 'theme' | string[];
  cardStyle?: 'outlined' | 'elevated' | 'flat';
}

export type DataSourceConfig = SpListSourceConfig | InlineSourceConfig | RestSourceConfig;

export interface SpListSourceConfig {
  type: 'spList';
  list?: string;
  listId?: string;
  webUrl?: string;
  select?: string[];
  expand?: string[];
  filter?: string;
  orderBy?: string;
  top?: number;
  pageSize?: number;
  rename?: Record<string, string>;
}

export interface InlineSourceConfig {
  type: 'inline';
  rows: Row[];
}

export interface RestSourceConfig {
  type: 'rest';
  url: string;
  path?: string;
  headers?: Record<string, string>;
  sharepoint?: boolean;
}

export interface DatasetConfig {
  from: string | string[];
  join?: JoinConfig[];
  filter?: string;
  compute?: Record<string, string>;
  groupBy?: Array<string | GroupByConfig>;
  aggregate?: Record<string, string>;
  having?: string;
  sort?: string | string[];
  limit?: number;
}

export interface JoinConfig {
  source: string;
  on: [string, string];
  as?: string;
  type?: 'left' | 'inner';
  many?: boolean;
}

export interface GroupByConfig {
  field: string;
  as?: string;
  bucket?: DateBucket;
}

export type DateBucket = 'day' | 'week' | 'month' | 'quarter' | 'year';

export type DatasetRef = string | DatasetConfig;

export interface WidgetBase {
  title?: string;
  subtitle?: string;
  span?: number;
  height?: number;
  dataset: DatasetRef;
}

export type WidgetConfig = KpiWidgetConfig | ChartWidgetConfig | TableWidgetConfig | TextWidgetConfig;

export type FormatName = 'number' | 'integer' | 'decimal' | 'percent' | 'currency' | 'compact' | 'date' | 'datetime' | 'text';

export interface FormatSpec {
  type?: FormatName;
  decimals?: number;
  currency?: string;
  prefix?: string;
  suffix?: string;
}

export type Format = FormatName | FormatSpec;

export type Level = 'good' | 'warning' | 'serious' | 'critical' | 'neutral' | 'info';

export interface ThresholdConfig {
  when: string;
  level: Level;
  label?: string;
}

export interface KpiWidgetConfig extends WidgetBase {
  type: 'kpi';
  value: string;
  label?: string;
  format?: Format;
  compare?: {
    value: string;
    label?: string;
    mode?: 'percent' | 'absolute';
    goodWhen?: 'up' | 'down';
    dataset?: DatasetRef;
  };
  target?: number | string;
  thresholds?: ThresholdConfig[];
  sparkline?: { dataset?: DatasetRef; x: string; y: string };
}

export type ChartType =
  | 'bar' | 'column' | 'line' | 'area' | 'pie' | 'donut' | 'scatter'
  | 'gauge' | 'funnel' | 'heatmap' | 'sankey' | 'treemap' | 'radar';

export interface SeriesConfig {
  field: string;
  name?: string;
  type?: 'bar' | 'line' | 'area';
  color?: string;
  stack?: string;
}

export interface ChartWidgetConfig extends WidgetBase {
  type: 'chart';
  chart: ChartType;
  x?: string;
  y?: string | Array<string | SeriesConfig>;
  series?: string;
  stacked?: boolean;
  from?: string;
  to?: string;
  value?: string;
  min?: number;
  max?: number;
  format?: Format;
  showLegend?: boolean;
  labels?: boolean | 'auto';
  smooth?: boolean;
  colors?: string[];
  echarts?: Record<string, unknown>;
}

export interface ColumnConfig {
  field: string;
  title?: string;
  format?: Format;
  width?: number;
  align?: 'left' | 'right' | 'center';
  bar?: boolean;
  link?: string;
  status?: ThresholdConfig[];
}

export interface TableWidgetConfig extends WidgetBase {
  type: 'table';
  columns?: Array<string | ColumnConfig>;
  pageSize?: number;
  search?: boolean;
  detail?: {
    dataset: DatasetRef;
    on: [string, string];
    title?: string;
    columns?: Array<string | ColumnConfig>;
  };
}

export interface TextWidgetConfig extends Omit<WidgetBase, 'dataset'> {
  type: 'text';
  text: string;
  dataset?: DatasetRef;
}

export interface ListViewConfig {
  source: string;
  title?: string;
  keyPrefix?: string;
  columns?: Array<string | ColumnConfig>;
  sort?: string | string[];
  filter?: string;
  pageSize?: number;
  search?: boolean;
  quickFilters?: string[];
  statusColors?: Record<string, Level>;
  form?: ItemFormConfig;
  autotag?: AutotagConfig;
}

export interface FormSection {
  title?: string;
  fields: string[];
  columns?: 1 | 2;
}

export type FormEntry = string | FormSection;

export interface ItemFormConfig {
  title?: string;
  description?: string;
  status?: string;
  main?: FormEntry[];
  side?: FormEntry[];
  fields?: string[];
  sideWidth?: number;
  readOnly?: boolean | string[];
  comments?: boolean;
}

export interface AiApiConfig {
  baseUrl: string;
  embeddingModel: string;
  chatModel?: string;
  batchSize?: number;
}

export interface TagDefinition {
  name: string;
  description?: string;
}

export interface AutotagConfig {
  field: string;
  api: AiApiConfig;
  tags?: TagDefinition[];
  tagDescriptions?: Record<string, string>;
  sourceFields?: string[];
  includeFileText?: boolean;
  maxChars?: number;
  maxTags?: number;
  minScore?: number;
  maxItems?: number;
  mode?: 'review' | 'auto';
  merge?: boolean;
  allow?: string[];
}

export interface DataViewSource {
  name: string;
  title?: string;
  columns?: Array<string | ColumnConfig>;
}

export interface DataViewConfig {
  title?: string;
  sources?: Array<string | DataViewSource>;
  pageSize?: number;
  export?: boolean;
}
