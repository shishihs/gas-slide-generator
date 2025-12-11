import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class KPIDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return requests;

        const cols = items.length > 4 ? 4 : (items.length || 1);
        const gap = layout.pxToPt(40);
        const cardW = (area.width - (gap * (cols - 1))) / cols;
        const cardH = layout.pxToPt(180);
        const y = area.top + (area.height - cardH) / 2 + layout.pxToPt(20);

        items.forEach((item: any, i: number) => {
            const x = area.left + i * (cardW + gap);
            const baseId = slideId + `_KPI_${i}`;

            // Divider Line
            if (i > 0) {
                const lineH = layout.pxToPt(100);
                const lineY = y + (cardH - lineH) / 2;
                const lineX = x - gap / 2;
                const lineId = baseId + '_LN';

                requests.push(RequestFactory.createLine(slideId, lineId, lineX, lineY, lineX, lineY + lineH));
                requests.push({
                    updateLineProperties: {
                        objectId: lineId,
                        lineProperties: {
                            lineFill: { solidFill: { color: { rgbColor: { red: 0.9, green: 0.9, blue: 0.95 } } } }, // Ghost Gray approx
                            weight: { magnitude: 1, unit: 'PT' }
                        },
                        fields: 'lineFill,weight'
                    }
                });
            }

            // Label
            const labelH = layout.pxToPt(20);
            const labelId = baseId + '_LBL';
            requests.push(RequestFactory.createShape(slideId, labelId, 'TEXT_BOX', x, y + layout.pxToPt(5), cardW, labelH));
            requests.push(RequestFactory.updateShapeProperties(labelId, null, null, null, 'BOTTOM'));
            requests.push(...BatchTextStyleUtils.setText(slideId, labelId, (item.label || 'METRIC').toUpperCase(), {
                size: 11, color: theme.colors.neutralGray, align: 'CENTER', bold: true
            }, theme));

            // Value
            const valueH = layout.pxToPt(90);
            const valueId = baseId + '_VAL';
            requests.push(RequestFactory.createShape(slideId, valueId, 'TEXT_BOX', x, y + labelH, cardW, valueH));
            requests.push(RequestFactory.updateShapeProperties(valueId, null, null, null, 'TOP'));

            const valStr = String(item.value || '0');
            let fontSize = 72;
            if (valStr.length > 4) fontSize = 60;
            if (valStr.length > 6) fontSize = 48;
            if (valStr.length > 10) fontSize = 36;

            requests.push(...BatchTextStyleUtils.setText(slideId, valueId, valStr, {
                size: fontSize, bold: true, color: settings.primaryColor, align: 'CENTER', fontType: 'lato'
            }, theme));

            // Status
            if (item.change || item.status) {
                const statusH = layout.pxToPt(30);
                const statusId = baseId + '_STS';
                requests.push(RequestFactory.createShape(slideId, statusId, 'TEXT_BOX', x, y + labelH + valueH, cardW, statusH));

                let color = theme.colors.neutralGray;
                let prefix = '';
                if (item.status === 'good') { color = '#2E7D32'; prefix = '↑ '; }
                if (item.status === 'bad') { color = '#C62828'; prefix = '↓ '; }

                requests.push(...BatchTextStyleUtils.setText(slideId, statusId, prefix + (item.change || ''), {
                    size: 14, bold: true, color: color, align: 'CENTER'
                }, theme));
            }
        });

        return requests;
    }
}
