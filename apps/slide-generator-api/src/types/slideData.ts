/**
 * SlideData Type Definitions
 * Generated from the slide generation prompt specification
 */

// =============================================================================
// Common Types
// =============================================================================

/** Inline markup allowed in text content */
// **太字** → Bold text
// [[重要語]] → Bold + Primary color (only in body columns, not headers)

/** Base properties shared by all slide types */
interface SlideBase {
    /** Speaker notes draft for the slide (plain text, no markup) */
    notes?: string;
}

// =============================================================================
// Slide Types
// =============================================================================

/** Title slide (表紙) */
export interface TitleSlide extends SlideBase {
    type: 'title';
    /** Main title (max 35 full-width chars) */
    title: string;
    /** Date in YYYY.MM.DD format */
    date: string;
}

/** Section divider slide (章扉) */
export interface SectionSlide extends SlideBase {
    type: 'section';
    /** Section title (max 30 full-width chars) */
    title: string;
    /** Section number (auto-numbered if omitted) */
    sectionNo?: number;
}

/** Closing slide (結び) */
export interface ClosingSlide extends SlideBase {
    type: 'closing';
}

/** Content slide - 1 or 2 columns with optional subhead */
export interface ContentSlide extends SlideBase {
    type: 'content';
    /** Slide title (max 40 full-width chars) */
    title: string;
    /** Subtitle (max 50 full-width chars, max 2 lines) */
    subhead?: string;
    /** Single column bullet points */
    points?: string[];
    /** Enable two-column layout */
    twoColumn?: boolean;
    /** Two-column content [leftColumn, rightColumn] */
    columns?: [string[], string[]];
}

/** Agenda slide - numbered agenda items */
export interface AgendaSlide extends SlideBase {
    type: 'agenda';
    title: string;
    subhead?: string;
    /** Agenda items (no numbers in text - auto-drawn) */
    items: string[];
}

/** Compare slide - left/right comparison */
export interface CompareSlide extends SlideBase {
    type: 'compare';
    title: string;
    subhead?: string;
    /** Left column header */
    leftTitle: string;
    /** Right column header */
    rightTitle: string;
    /** Left column items */
    leftItems: string[];
    /** Right column items */
    rightItems: string[];
}

/** Process slide - visual step format (max 4 steps) */
export interface ProcessSlide extends SlideBase {
    type: 'process';
    title: string;
    subhead?: string;
    /** Process steps (no step numbers - auto-drawn) */
    steps: string[];
}

/** Process list slide - simple list format */
export interface ProcessListSlide extends SlideBase {
    type: 'processList';
    title: string;
    subhead?: string;
    /** Process steps (no step numbers - auto-drawn) */
    steps: string[];
}

/** Timeline milestone state */
export type MilestoneState = 'done' | 'next' | 'todo';

/** Timeline milestone item */
export interface TimelineMilestone {
    /** Milestone label (max 30 chars, concise phase name recommended) */
    label: string;
    /** Date or time label */
    date: string;
    /** Milestone state */
    state?: MilestoneState;
}

/** Timeline slide - chronological milestones */
export interface TimelineSlide extends SlideBase {
    type: 'timeline';
    title: string;
    subhead?: string;
    /** Timeline milestones (no order prefixes - auto-drawn) */
    milestones: TimelineMilestone[];
}

/** Diagram lane item */
export interface DiagramLane {
    /** Lane title */
    title: string;
    /** Lane items */
    items: string[];
}

/** Diagram slide - swim lane diagram */
export interface DiagramSlide extends SlideBase {
    type: 'diagram';
    title: string;
    subhead?: string;
    /** Swim lanes */
    lanes: DiagramLane[];
}

/** Cycle item */
export interface CycleItem {
    /** Main label */
    label: string;
    /** Secondary label */
    subLabel?: string;
}

/** Cycle slide - circular diagram (exactly 4 items) */
export interface CycleSlide extends SlideBase {
    type: 'cycle';
    title: string;
    subhead?: string;
    /** Cycle items (exactly 4 items, ~20 chars each) */
    items: CycleItem[];
    /** Center text */
    centerText?: string;
}

/** Card item - simple or with title/desc */
export type CardItem = string | { title: string; desc?: string };

/** Cards slide - simple cards (max 6 items, 3x2) */
export interface CardsSlide extends SlideBase {
    type: 'cards';
    title: string;
    subhead?: string;
    /** Number of columns */
    columns?: 2 | 3;
    /** Card items */
    items: CardItem[];
}

/** Header card item */
export interface HeaderCardItem {
    /** Card header (white text on colored background) */
    title: string;
    /** Card description */
    desc?: string;
}

/** Header cards slide - cards with colored headers (max 6 items) */
export interface HeaderCardsSlide extends SlideBase {
    type: 'headerCards';
    title: string;
    subhead?: string;
    columns?: 2 | 3;
    items: HeaderCardItem[];
}

/** Table slide */
export interface TableSlide extends SlideBase {
    type: 'table';
    title: string;
    subhead?: string;
    /** Table headers */
    headers: string[];
    /** Table rows */
    rows: string[][];
}

/** Progress item */
export interface ProgressItem {
    /** Progress label */
    label: string;
    /** Progress percentage (0-100) */
    percent: number;
}

/** Progress slide - progress bars */
export interface ProgressSlide extends SlideBase {
    type: 'progress';
    title: string;
    subhead?: string;
    items: ProgressItem[];
}

/** Quote slide - quotation with attribution */
export interface QuoteSlide extends SlideBase {
    type: 'quote';
    title: string;
    subhead?: string;
    /** Quote text */
    text: string;
    /** Quote author/source */
    author: string;
}

/** KPI status indicator */
export type KpiStatus = 'good' | 'bad' | 'neutral';

/** KPI item */
export interface KpiItem {
    /** KPI label */
    label: string;
    /** KPI value (formatted string) */
    value: string;
    /** Change indicator (e.g., "+5%", "-10%") */
    change: string;
    /** Status color */
    status: KpiStatus;
}

/** KPI slide - key performance indicators (max 4 items) */
export interface KpiSlide extends SlideBase {
    type: 'kpi';
    title: string;
    subhead?: string;
    /** Number of columns */
    columns?: 2 | 3 | 4;
    items: KpiItem[];
}

/** Bullet card item */
export interface BulletCardItem {
    /** Card title */
    title: string;
    /** Card description */
    desc: string;
}

/** Bullet cards slide - cards with bullet points (max 3 items) */
export interface BulletCardsSlide extends SlideBase {
    type: 'bulletCards';
    title: string;
    subhead?: string;
    items: BulletCardItem[];
}

/** FAQ item */
export interface FaqItem {
    /** Question (max 28 chars) */
    q: string;
    /** Answer (max 45 chars) */
    a: string;
}

/** FAQ slide - question and answer format (1-4 items) */
export interface FaqSlide extends SlideBase {
    type: 'faq';
    title: string;
    subhead?: string;
    items: FaqItem[];
}

/** Trend direction */
export type TrendDirection = 'up' | 'down' | 'neutral';

/** Stats compare item */
export interface StatsCompareItem {
    /** Comparison label (max 12 chars) */
    label: string;
    /** Left value */
    leftValue: string;
    /** Right value */
    rightValue: string;
    /** Trend indicator */
    trend?: TrendDirection;
}

/** Stats compare slide - numerical comparison table */
export interface StatsCompareSlide extends SlideBase {
    type: 'statsCompare';
    title: string;
    subhead?: string;
    /** Left column header */
    leftTitle: string;
    /** Right column header */
    rightTitle: string;
    stats: StatsCompareItem[];
}

/** Bar compare item */
export interface BarCompareItem {
    /** Comparison label (max 12 chars) */
    label: string;
    /** Left value */
    leftValue: string;
    /** Right value */
    rightValue: string;
    /** Trend indicator */
    trend?: TrendDirection;
}

/** Bar compare slide - horizontal bar comparison */
export interface BarCompareSlide extends SlideBase {
    type: 'barCompare';
    title: string;
    subhead?: string;
    stats: BarCompareItem[];
    /** Show trend indicators (default: false) */
    showTrends?: boolean;
}

/** Triangle item */
export interface TriangleItem {
    /** Keyword/short phrase (10-12 chars max) */
    title: string;
    /** Brief description (15 chars max) */
    desc?: string;
}

/** Triangle slide - 3-point triangle diagram (exactly 3 items) */
export interface TriangleSlide extends SlideBase {
    type: 'triangle';
    title: string;
    subhead?: string;
    /** Exactly 3 items */
    items: [TriangleItem, TriangleItem, TriangleItem];
}

/** Pyramid level */
export interface PyramidLevel {
    /** Level title */
    title: string;
    /** Level description */
    description: string;
}

/** Pyramid slide - hierarchical pyramid (3-4 levels) */
export interface PyramidSlide extends SlideBase {
    type: 'pyramid';
    title: string;
    subhead?: string;
    /** Pyramid levels (3-4 levels) */
    levels: PyramidLevel[];
}

/** Flow chart flow */
export interface FlowChartFlow {
    /** Steps in this flow (2-4 steps per row) */
    steps: string[];
}

/** Flow chart slide - left-to-right flow (1-2 rows, max 8 total) */
export interface FlowChartSlide extends SlideBase {
    type: 'flowChart';
    title: string;
    subhead?: string;
    /** Flow rows (1-2 flows) */
    flows: FlowChartFlow[];
}

/** Step up item */
export interface StepUpItem {
    /** Step title (max 10 chars) */
    title: string;
    /** Step description (max 28 chars) */
    desc: string;
}

/** Step up slide - stair-step growth visualization (2-5 steps) */
export interface StepUpSlide extends SlideBase {
    type: 'stepUp';
    title: string;
    subhead?: string;
    items: StepUpItem[];
}

/** Image position */
export type ImagePosition = 'left' | 'right';

/** Image text slide - image with text columns */
export interface ImageTextSlide extends SlideBase {
    type: 'imageText';
    title: string;
    subhead?: string;
    /** 
     * Image source:
     * - URL starting with http:// or https://
     * - Video timestamp in mm:ss.S format
     * - Image description (30 chars) for attached images
     * - Chart configuration object
     */
    image: string | ChartData;
    /** Image caption */
    imageCaption?: string;
    /** Image position */
    imagePosition?: ImagePosition;
    /** Text bullet points */
    points: string[];
}

// =============================================================================
// Chart Types
// =============================================================================

/** Gradient color definition */
export interface GradientColor {
    /** Gradient start color (hex) */
    start: string;
    /** Gradient end color (hex) */
    end: string;
}

/** Series color with ID */
export interface SeriesColor extends GradientColor {
    /** Series identifier */
    id: string;
}

/** Chart layout configuration */
export interface ChartLayout {
    width: number;
    height: number;
    marginTop: number;
    marginBottom: number;
    marginLeft: number;
    marginRight: number;
    horizontalPadding?: number;
}

/** Y-axis configuration */
export interface YAxisConfig {
    min?: number;
    max?: number;
    tickCount?: number;
    unit?: string;
}

// -----------------------------------------------------------------------------
// Multi-Line Chart
// -----------------------------------------------------------------------------

/** Multi-line chart series */
export interface MultiLineSeries {
    id: string;
    label: string;
    values: number[];
}

/** Line chart options */
export interface LineOptions {
    markerRadius?: number;
    dataLabelOffsetY?: number;
    horizontalPadding?: number;
}

/** Multi-line chart data */
export interface MultiLineChartData {
    title: string;
    subtitle: string;
    source: string;
    yAxisUnitLabel: string;
    xAxisLabels: string[];
    series: MultiLineSeries[];
    colors: SeriesColor[];
    layout: ChartLayout;
    yAxis: YAxisConfig;
    lineOptions?: LineOptions;
}

/** Multi-line chart */
export interface MultiLineChart {
    chartType: 'multi-line';
    data: MultiLineChartData;
}

// -----------------------------------------------------------------------------
// Donut Chart
// -----------------------------------------------------------------------------

/** Donut chart item */
export interface DonutItem {
    label: string;
    value: number;
    id: string;
}

/** Donut chart data */
export interface DonutChartData {
    title: string;
    subtitle: string;
    source: string;
    centerLabel: string;
    colors: SeriesColor[];
    items: DonutItem[];
}

/** Donut chart */
export interface DonutChart {
    chartType: 'donut';
    data: DonutChartData;
}

// -----------------------------------------------------------------------------
// Stacked Bar Chart
// -----------------------------------------------------------------------------

/** Bar data item */
export interface BarDataItem {
    label: string;
    /** Values must be >= 0 */
    values: number[];
}

/** Bar options */
export interface BarOptions {
    width?: number;
    cornerRadius?: number;
    totalLabelOffset?: number;
    barToSlotRatio?: number;
    labelPosition?: 'auto' | string;
}

/** Stacked bar chart data */
export interface StackedBarChartData {
    title: string;
    subtitle: string;
    source: string;
    yAxisUnitLabel: string;
    colors: SeriesColor[];
    legendLabels: string[];
    barData: BarDataItem[];
    layout: ChartLayout;
    barOptions?: BarOptions;
    yAxis: YAxisConfig;
}

/** Stacked bar chart (absolute values) */
export interface StackedBarChart {
    chartType: 'stacked-bar';
    data: StackedBarChartData;
}

// -----------------------------------------------------------------------------
// 100% Stacked Bar Chart
// -----------------------------------------------------------------------------

/** 100% stacked bar chart data */
export interface StackedBar100ChartData {
    title: string;
    subtitle: string;
    source: string;
    colors: SeriesColor[];
    legendLabels: string[];
    barData: BarDataItem[];
    layout: ChartLayout;
    barOptions?: BarOptions;
    yAxis: Pick<YAxisConfig, 'tickCount'>;
}

/** 100% Stacked bar chart */
export interface StackedBar100Chart {
    chartType: '100-stacked-bar';
    data: StackedBar100ChartData;
}

// -----------------------------------------------------------------------------
// Bar Chart (Single Series)
// -----------------------------------------------------------------------------

/** Bar chart item */
export interface BarChartItem {
    label: string;
    value: number;
}

/** Bar chart data */
export interface BarChartData {
    title: string;
    subtitle: string;
    source: string;
    items: BarChartItem[];
    color: GradientColor;
    layout: ChartLayout;
    barOptions?: Pick<BarOptions, 'barToSlotRatio'>;
    yAxis: YAxisConfig;
}

/** Bar chart (single series) */
export interface BarChart {
    chartType: 'bar';
    data: BarChartData;
}

// -----------------------------------------------------------------------------
// Combo Chart (Bar + Line)
// -----------------------------------------------------------------------------

/** Combo chart item */
export interface ComboChartItem {
    label: string;
    barValue: number;
    lineValue: number;
}

/** Combo chart colors */
export interface ComboChartColors {
    bar: GradientColor;
    line: string;
}

/** Combo chart data */
export interface ComboChartData {
    title: string;
    subtitle: string;
    source: string;
    legendBarLabel: string;
    legendLineLabel: string;
    yAxisLeftLabel: string;
    yAxisRightLabel: string;
    colors: ComboChartColors;
    items: ComboChartItem[];
    layout: ChartLayout;
    barOptions?: BarOptions;
    lineOptions?: Pick<LineOptions, 'markerRadius'>;
    yAxisLeft: YAxisConfig;
    yAxisRight: YAxisConfig;
}

/** Combo chart (bar + line) */
export interface ComboChart {
    chartType: 'combo';
    data: ComboChartData;
}

// -----------------------------------------------------------------------------
// Line Chart (Single Series)
// -----------------------------------------------------------------------------

/** Line chart item */
export interface LineChartItem {
    label: string;
    value: number;
}

/** Line chart color */
export interface LineChartColor extends GradientColor {
    line: string;
    label: string;
}

/** Line chart data */
export interface LineChartData {
    title: string;
    subtitle: string;
    source: string;
    yAxisUnitLabel: string;
    items: LineChartItem[];
    color: LineChartColor;
    layout: ChartLayout;
    yAxis: YAxisConfig;
    lineOptions?: LineOptions;
}

/** Line chart (single series) */
export interface LineChart {
    chartType: 'line';
    data: LineChartData;
}

// -----------------------------------------------------------------------------
// Chart Union Type
// -----------------------------------------------------------------------------

/** All chart types */
export type ChartData =
    | MultiLineChart
    | DonutChart
    | StackedBarChart
    | StackedBar100Chart
    | BarChart
    | ComboChart
    | LineChart;

// =============================================================================
// Slide Union Type
// =============================================================================

/** All slide types */
export type Slide =
    | TitleSlide
    | SectionSlide
    | ClosingSlide
    | ContentSlide
    | AgendaSlide
    | CompareSlide
    | ProcessSlide
    | ProcessListSlide
    | TimelineSlide
    | DiagramSlide
    | CycleSlide
    | CardsSlide
    | HeaderCardsSlide
    | TableSlide
    | ProgressSlide
    | QuoteSlide
    | KpiSlide
    | BulletCardsSlide
    | FaqSlide
    | StatsCompareSlide
    | BarCompareSlide
    | TriangleSlide
    | PyramidSlide
    | FlowChartSlide
    | StepUpSlide
    | ImageTextSlide;

/** The main slideData array type */
export type SlideData = Slide[];

// =============================================================================
// Type Guards
// =============================================================================

/** Check if slide is a title slide */
export function isTitleSlide(slide: Slide): slide is TitleSlide {
    return slide.type === 'title';
}

/** Check if slide is a section slide */
export function isSectionSlide(slide: Slide): slide is SectionSlide {
    return slide.type === 'section';
}

/** Check if slide is a closing slide */
export function isClosingSlide(slide: Slide): slide is ClosingSlide {
    return slide.type === 'closing';
}

/** Check if slide is a content slide */
export function isContentSlide(slide: Slide): slide is ContentSlide {
    return slide.type === 'content';
}

/** Check if slide is an agenda slide */
export function isAgendaSlide(slide: Slide): slide is AgendaSlide {
    return slide.type === 'agenda';
}

/** Check if slide is a compare slide */
export function isCompareSlide(slide: Slide): slide is CompareSlide {
    return slide.type === 'compare';
}

/** Check if slide is a process slide */
export function isProcessSlide(slide: Slide): slide is ProcessSlide {
    return slide.type === 'process';
}

/** Check if slide is a process list slide */
export function isProcessListSlide(slide: Slide): slide is ProcessListSlide {
    return slide.type === 'processList';
}

/** Check if slide is a timeline slide */
export function isTimelineSlide(slide: Slide): slide is TimelineSlide {
    return slide.type === 'timeline';
}

/** Check if slide is a diagram slide */
export function isDiagramSlide(slide: Slide): slide is DiagramSlide {
    return slide.type === 'diagram';
}

/** Check if slide is a cycle slide */
export function isCycleSlide(slide: Slide): slide is CycleSlide {
    return slide.type === 'cycle';
}

/** Check if slide is a cards slide */
export function isCardsSlide(slide: Slide): slide is CardsSlide {
    return slide.type === 'cards';
}

/** Check if slide is a header cards slide */
export function isHeaderCardsSlide(slide: Slide): slide is HeaderCardsSlide {
    return slide.type === 'headerCards';
}

/** Check if slide is a table slide */
export function isTableSlide(slide: Slide): slide is TableSlide {
    return slide.type === 'table';
}

/** Check if slide is a progress slide */
export function isProgressSlide(slide: Slide): slide is ProgressSlide {
    return slide.type === 'progress';
}

/** Check if slide is a quote slide */
export function isQuoteSlide(slide: Slide): slide is QuoteSlide {
    return slide.type === 'quote';
}

/** Check if slide is a KPI slide */
export function isKpiSlide(slide: Slide): slide is KpiSlide {
    return slide.type === 'kpi';
}

/** Check if slide is a bullet cards slide */
export function isBulletCardsSlide(slide: Slide): slide is BulletCardsSlide {
    return slide.type === 'bulletCards';
}

/** Check if slide is a FAQ slide */
export function isFaqSlide(slide: Slide): slide is FaqSlide {
    return slide.type === 'faq';
}

/** Check if slide is a stats compare slide */
export function isStatsCompareSlide(slide: Slide): slide is StatsCompareSlide {
    return slide.type === 'statsCompare';
}

/** Check if slide is a bar compare slide */
export function isBarCompareSlide(slide: Slide): slide is BarCompareSlide {
    return slide.type === 'barCompare';
}

/** Check if slide is a triangle slide */
export function isTriangleSlide(slide: Slide): slide is TriangleSlide {
    return slide.type === 'triangle';
}

/** Check if slide is a pyramid slide */
export function isPyramidSlide(slide: Slide): slide is PyramidSlide {
    return slide.type === 'pyramid';
}

/** Check if slide is a flow chart slide */
export function isFlowChartSlide(slide: Slide): slide is FlowChartSlide {
    return slide.type === 'flowChart';
}

/** Check if slide is a step up slide */
export function isStepUpSlide(slide: Slide): slide is StepUpSlide {
    return slide.type === 'stepUp';
}

/** Check if slide is an image text slide */
export function isImageTextSlide(slide: Slide): slide is ImageTextSlide {
    return slide.type === 'imageText';
}

/** Check if a value is a chart data object */
export function isChartData(value: unknown): value is ChartData {
    return (
        typeof value === 'object' &&
        value !== null &&
        'chartType' in value &&
        'data' in value
    );
}
