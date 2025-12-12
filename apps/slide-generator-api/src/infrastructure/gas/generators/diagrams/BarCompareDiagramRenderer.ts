import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';
import { generateCompareColors } from '../../../../common/utils/ColorUtils';

export class BarCompareDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const leftTitle = data.leftTitle || '導入前';
        const rightTitle = data.rightTitle || '導入後';
        const stats = data.stats || [];
        if (!stats.length) return requests;

        const compareColors = generateCompareColors(settings.primaryColor);

        // Find max value for scaling
        let maxValue = 0;
        stats.forEach((stat: any) => {
            const leftNum = parseFloat(String(stat.leftValue || '0').replace(/[^0-9.]/g, '')) || 0;
            const rightNum = parseFloat(String(stat.rightValue || '0').replace(/[^0-9.]/g, '')) || 0;
            maxValue = Math.max(maxValue, leftNum, rightNum);
        });
        if (maxValue === 0) maxValue = 100;

        // Layout
        const labelColW = area.width * 0.2;
        const barAreaW = area.width * 0.6;
        // const valueColW = area.width * 0.1; // Unused in original
        // const trendColW = area.width * 0.1; // Unused in original

        const rowHeight = Math.min(layout.pxToPt(80), area.height / stats.length);
        const barHeight = layout.pxToPt(18);
        let currentY = area.top;



        stats.forEach((stat: any, index: number) => {
            const label = stat.label || '';
            const leftValue = stat.leftValue || '';
            const rightValue = stat.rightValue || '';
            const trend = stat.trend || null;

            const leftNum = parseFloat(String(leftValue).replace(/[^0-9.]/g, '')) || 0;
            const rightNum = parseFloat(String(rightValue).replace(/[^0-9.]/g, '')) || 0;

            const baseId = slideId + '_BAR_R' + index;

            // Label (Left Column)
            const labelId = baseId + '_LBL';
            requests.push(RequestFactory.createShape(slideId, labelId, 'TEXT_BOX', area.left, currentY, labelColW, rowHeight));
            requests.push(...BatchTextStyleUtils.setText(slideId, labelId, label, {
                size: 14, bold: true, align: 'START', color: theme.colors.textPrimary
            }, theme));
            requests.push(RequestFactory.updateShapeProperties(labelId, null, null, null, 'MIDDLE'));

            // Bar Area
            const barLeft = area.left + labelColW + layout.pxToPt(20);
            const maxBarWidth = area.width - (labelColW + layout.pxToPt(80));

            const barGap = layout.pxToPt(2);
            const barY = currentY + (rowHeight - (barHeight * 2 + barGap * 2)) / 2;

            // Left Bar (Gray)
            const leftBarW = (leftNum / maxValue) * maxBarWidth;
            if (leftBarW > 0) {
                const lbId = baseId + '_LB';
                requests.push(RequestFactory.createShape(slideId, lbId, 'RECTANGLE', barLeft, barY, leftBarW, barHeight));
                requests.push(RequestFactory.updateShapeProperties(lbId, theme.colors.neutralGray, 'TRANSPARENT'));

                // Value Text (End)
                const lbvId = baseId + '_LBV';
                requests.push(RequestFactory.createShape(slideId, lbvId, 'TEXT_BOX', barLeft + leftBarW + layout.pxToPt(5), barY - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5)));
                requests.push(...BatchTextStyleUtils.setText(slideId, lbvId, leftValue, { size: 10, color: theme.colors.neutralGray }, theme));
            }

            // Right Bar (Color)
            const rightBarW = (rightNum / maxValue) * maxBarWidth;
            if (rightBarW > 0) {
                const rbId = baseId + '_RB';
                requests.push(RequestFactory.createShape(slideId, rbId, 'RECTANGLE', barLeft, barY + barHeight + barGap, rightBarW, barHeight));
                requests.push(RequestFactory.updateShapeProperties(rbId, settings.primaryColor, 'TRANSPARENT'));

                const rbvId = baseId + '_RBV';
                requests.push(RequestFactory.createShape(slideId, rbvId, 'TEXT_BOX', barLeft + rightBarW + layout.pxToPt(5), barY + barHeight + barGap - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5)));
                requests.push(...BatchTextStyleUtils.setText(slideId, rbvId, rightValue, { size: 10, bold: true, color: settings.primaryColor }, theme));
            }

            // Trend Arrow
            if (trend) {
                const trendX = area.left + area.width - layout.pxToPt(40);
                const isUp = String(trend).toLowerCase() === 'up';
                const trId = baseId + '_TR';
                requests.push(RequestFactory.createShape(slideId, trId, 'TEXT_BOX', trendX, currentY, layout.pxToPt(40), rowHeight));

                const color = isUp ? '#2E7D32' : '#C62828';
                const sym = isUp ? '↑' : '↓';
                requests.push(...BatchTextStyleUtils.setText(slideId, trId, sym, { size: 20, bold: true, color: color, align: 'CENTER' }, theme));
                requests.push(RequestFactory.updateShapeProperties(trId, null, null, null, 'MIDDLE'));
            }

            // Separator Line
            if (index < stats.length - 1) {
                const lineY = currentY + rowHeight;
                const lineId = baseId + '_LN';
                requests.push(RequestFactory.createLine(slideId, lineId, area.left, lineY, area.left + area.width, lineY));

                // Manual updateLineProperties because RequestFactory doesn't have it yet
                requests.push({
                    updateLineProperties: {
                        objectId: lineId,
                        lineProperties: {
                            lineFill: { solidFill: { color: { rgbColor: RequestFactory.toRgbColor(theme.colors.ghostGray) || { red: 0.9, green: 0.9, blue: 0.9 } } } },
                            weight: { magnitude: 0.5, unit: 'PT' }
                        },
                        fields: 'lineFill,weight'
                    }
                });
            }

            currentY += rowHeight;
        });

        return requests;
    }
}
