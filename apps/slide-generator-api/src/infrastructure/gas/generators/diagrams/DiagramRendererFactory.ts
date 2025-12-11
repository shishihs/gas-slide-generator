import { IDiagramRenderer } from './IDiagramRenderer';
import { CardsDiagramRenderer } from './CardsDiagramRenderer';
import { TimelineDiagramRenderer } from './TimelineDiagramRenderer';
import { ProcessDiagramRenderer } from './ProcessDiagramRenderer';
import { CycleDiagramRenderer } from './CycleDiagramRenderer';
import { PyramidDiagramRenderer } from './PyramidDiagramRenderer';
import { TriangleDiagramRenderer } from './TriangleDiagramRenderer';
import { ComparisonDiagramRenderer } from './ComparisonDiagramRenderer';
import { StatsCompareDiagramRenderer } from './StatsCompareDiagramRenderer';
import { BarCompareDiagramRenderer } from './BarCompareDiagramRenderer';
import { StepUpDiagramRenderer } from './StepUpDiagramRenderer';
import { LanesDiagramRenderer } from './LanesDiagramRenderer';
import { FlowChartDiagramRenderer } from './FlowChartDiagramRenderer';
import { KPIDiagramRenderer } from './KPIDiagramRenderer';
import { TableDiagramRenderer } from './TableDiagramRenderer';
import { FAQDiagramRenderer } from './FAQDiagramRenderer';
import { QuoteDiagramRenderer } from './QuoteDiagramRenderer';
import { ProgressDiagramRenderer } from './ProgressDiagramRenderer';
import { ImageTextDiagramRenderer } from './ImageTextDiagramRenderer';

export class DiagramRendererFactory {
    static getRenderer(type: string): IDiagramRenderer | null {
        const normalizedType = type.toLowerCase();
        Logger.log(`[Factory] Checking renderer for type: '${type}' (normalized: '${normalizedType}')`);

        if (normalizedType.includes('cards') || normalizedType.includes('headercards')) {
            Logger.log('[Factory] Matched CardsDiagramRenderer');
            return new CardsDiagramRenderer();
        }
        if (normalizedType.includes('timeline')) {
            Logger.log('[Factory] Matched TimelineDiagramRenderer');
            return new TimelineDiagramRenderer();
        }
        if (normalizedType.includes('process')) {
            Logger.log('[Factory] Matched ProcessDiagramRenderer');
            return new ProcessDiagramRenderer();
        }
        if (normalizedType.includes('cycle')) {
            Logger.log('[Factory] Matched CycleDiagramRenderer');
            return new CycleDiagramRenderer();
        }
        if (normalizedType.includes('pyramid')) {
            Logger.log('[Factory] Matched PyramidDiagramRenderer');
            return new PyramidDiagramRenderer();
        }
        if (normalizedType.includes('triangle')) {
            Logger.log('[Factory] Matched TriangleDiagramRenderer');
            return new TriangleDiagramRenderer();
        }
        if (normalizedType.includes('compare') || normalizedType.includes('comparison') || normalizedType.includes('kaizen')) {
            // 'stats' -> StatsCompare, 'bar' -> BarCompare, else General Compare
            if (normalizedType.includes('stats')) {
                Logger.log('[Factory] Matched StatsCompareDiagramRenderer');
                return new StatsCompareDiagramRenderer();
            }
            if (normalizedType.includes('bar')) {
                Logger.log('[Factory] Matched BarCompareDiagramRenderer');
                return new BarCompareDiagramRenderer();
            }
            Logger.log('[Factory] Matched ComparisonDiagramRenderer');
            return new ComparisonDiagramRenderer();
        }
        if (normalizedType.includes('stepup') || normalizedType.includes('stair')) {
            Logger.log('[Factory] Matched StepUpDiagramRenderer');
            return new StepUpDiagramRenderer();
        }
        if (normalizedType.includes('lanes') || normalizedType.includes('diagram')) {
            Logger.log('[Factory] Matched LanesDiagramRenderer');
            return new LanesDiagramRenderer();
        }
        if (normalizedType.includes('flow')) { // flow or flowchart
            Logger.log('[Factory] Matched FlowChartDiagramRenderer');
            return new FlowChartDiagramRenderer();
        }
        if (normalizedType.includes('kpi')) {
            Logger.log('[Factory] Matched KPIDiagramRenderer');
            return new KPIDiagramRenderer();
        }
        if (normalizedType.includes('table')) {
            Logger.log('[Factory] Matched TableDiagramRenderer');
            return new TableDiagramRenderer();
        }
        if (normalizedType.includes('faq')) {
            Logger.log('[Factory] Matched FAQDiagramRenderer');
            return new FAQDiagramRenderer();
        }
        if (normalizedType.includes('quote')) {
            Logger.log('[Factory] Matched QuoteDiagramRenderer');
            return new QuoteDiagramRenderer();
        }
        if (normalizedType.includes('progress')) {
            Logger.log('[Factory] Matched ProgressDiagramRenderer');
            return new ProgressDiagramRenderer();
        }
        if (normalizedType.includes('image') || normalizedType.includes('imagetext')) {
            Logger.log('[Factory] Matched ImageTextDiagramRenderer');
            return new ImageTextDiagramRenderer();
        }

        Logger.log('[Factory] No matching renderer found.');
        return null;
    }
}

