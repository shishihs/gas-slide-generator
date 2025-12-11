import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class FlowChartDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const steps = data.steps || data.items || [];
        if (!steps.length) return requests;

        const count = steps.length;
        const gap = 30;
        const boxWidth = (area.width - (gap * (count - 1))) / count;
        const boxHeight = 80;
        const y = area.top + (area.height - boxHeight) / 2;

        for (let i = 0; i < count; i++) {
            const step = steps[i];
            const x = area.left + i * (boxWidth + gap);
            const baseId = slideId + `_FLOW_${i}`;
            const boxId = baseId + '_BOX';
            const arrowId = baseId + '_ARR';

            // 1. Box
            requests.push(RequestFactory.createShape(slideId, boxId, 'ROUND_RECTANGLE', x, y, boxWidth, boxHeight));
            requests.push(RequestFactory.updateShapeProperties(boxId, theme.colors.backgroundGray, settings.primaryColor, 2, 'MIDDLE'));

            const label = typeof step === 'string' ? step : (step.label || '');
            requests.push(...BatchTextStyleUtils.setText(slideId, boxId, label, {
                size: theme.fonts.sizes.body,
                align: 'CENTER',
                color: theme.colors.textPrimary
            }, theme));

            // 2. Connector Arrow (if not last)
            if (i < count - 1) {
                const ax = x + boxWidth;
                const ay = y + boxHeight / 2;
                const bx = x + boxWidth + gap;
                const by = ay;

                requests.push(RequestFactory.createLine(slideId, arrowId, ax, ay, bx, by));
                requests.push({
                    updateLineProperties: {
                        objectId: arrowId,
                        lineProperties: {
                            lineFill: { solidFill: { color: { rgbColor: { red: 0.6, green: 0.6, blue: 0.6 } } } }, // Neutral Gray (approx)
                            weight: { magnitude: 1, unit: 'PT' },
                            endArrow: 'ARROW'
                        },
                        fields: 'lineFill,weight,endArrow'
                    }
                });
            }
        }

        return requests;
    }
}
