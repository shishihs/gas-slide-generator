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

        if (normalizedType.includes('cards') || normalizedType.includes('headercards')) {
            return new CardsDiagramRenderer();
        }
        if (normalizedType.includes('timeline')) {
            return new TimelineDiagramRenderer();
        }
        if (normalizedType.includes('process')) {
            return new ProcessDiagramRenderer();
        }
        if (normalizedType.includes('cycle')) {
            return new CycleDiagramRenderer();
        }
        if (normalizedType.includes('pyramid')) {
            return new PyramidDiagramRenderer();
        }
        if (normalizedType.includes('triangle')) {
            return new TriangleDiagramRenderer();
        }
        if (normalizedType.includes('compare') || normalizedType.includes('comparison') || normalizedType.includes('kaizen')) {
            // 'stats' -> StatsCompare, 'bar' -> BarCompare, else General Compare
            if (normalizedType.includes('stats')) return new StatsCompareDiagramRenderer();
            if (normalizedType.includes('bar')) return new BarCompareDiagramRenderer();
            return new ComparisonDiagramRenderer();
        }
        if (normalizedType.includes('stepup') || normalizedType.includes('stair')) {
            return new StepUpDiagramRenderer();
        }
        if (normalizedType.includes('lanes') || normalizedType.includes('diagram')) {
            return new LanesDiagramRenderer();
        }
        if (normalizedType.includes('flow')) { // flow or flowchart
            return new FlowChartDiagramRenderer();
        }
        if (normalizedType.includes('kpi')) {
            return new KPIDiagramRenderer();
        }
        if (normalizedType.includes('table')) {
            return new TableDiagramRenderer();
        }
        if (normalizedType.includes('faq')) {
            return new FAQDiagramRenderer();
        }
        if (normalizedType.includes('quote')) {
            return new QuoteDiagramRenderer();
        }
        if (normalizedType.includes('progress')) {
            return new ProgressDiagramRenderer();
        }
        if (normalizedType.includes('image') || normalizedType.includes('imagetext')) {
            return new ImageTextDiagramRenderer();
        }

        return null;
    }
}

