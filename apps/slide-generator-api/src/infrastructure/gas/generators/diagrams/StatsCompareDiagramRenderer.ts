import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';
import { generateCompareColors } from '../../../../common/utils/ColorUtils';

export class StatsCompareDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const leftTitle = data.leftTitle || '導入前';
        const rightTitle = data.rightTitle || '導入後';
        const stats = data.stats || [];
        if (!stats.length) return requests;

        const compareColors = generateCompareColors(settings.primaryColor);

        // Header row
        const headerH = layout.pxToPt(45);
        const labelColW = area.width * 0.35;
        const valueColW = (area.width - labelColW) / 2;

        // Left Title Header
        const leftHeaderX = area.left + labelColW;
        const lhId = slideId + '_STAT_LH';
        requests.push(RequestFactory.createShape(slideId, lhId, 'RECTANGLE', leftHeaderX, area.top, valueColW, headerH));
        requests.push(RequestFactory.updateShapeProperties(lhId, compareColors.left, 'TRANSPARENT', 0, 'MIDDLE'));
        requests.push(...BatchTextStyleUtils.setText(slideId, lhId, leftTitle, { size: 14, bold: true, color: theme.colors.backgroundGray, align: 'CENTER' }, theme));

        // Right Title Header
        const rightHeaderX = area.left + labelColW + valueColW;
        const rhId = slideId + '_STAT_RH';
        requests.push(RequestFactory.createShape(slideId, rhId, 'RECTANGLE', rightHeaderX, area.top, valueColW, headerH));
        requests.push(RequestFactory.updateShapeProperties(rhId, compareColors.right, 'TRANSPARENT', 0, 'MIDDLE'));
        requests.push(...BatchTextStyleUtils.setText(slideId, rhId, rightTitle, { size: 14, bold: true, color: theme.colors.backgroundGray, align: 'CENTER' }, theme));

        // Data rows
        const availableHeight = area.height - headerH;
        const rowHeight = Math.min(layout.pxToPt(60), availableHeight / stats.length);
        let currentY = area.top + headerH;

        stats.forEach((stat: any, index: number) => {
            const label = stat.label || '';
            const leftValue = stat.leftValue || '';
            const rightValue = stat.rightValue || '';
            const trend = stat.trend || null;

            const rowBg = index % 2 === 0 ? theme.colors.backgroundGray : '#FFFFFF';

            // Unique IDs per row
            const baseId = slideId + '_STAT_R' + index;

            // Label cell
            const lId = baseId + '_L';
            requests.push(RequestFactory.createShape(slideId, lId, 'RECTANGLE', area.left, currentY, labelColW, rowHeight));
            requests.push(RequestFactory.updateShapeProperties(lId, rowBg, theme.colors.faintGray, 1, 'MIDDLE')); // Border width 1? Was 1 in mock? default
            requests.push(...BatchTextStyleUtils.setText(slideId, lId, label, { size: theme.fonts.sizes.body, bold: true, align: 'START' }, theme));

            // Left value cell
            const lvId = baseId + '_LV';
            requests.push(RequestFactory.createShape(slideId, lvId, 'RECTANGLE', leftHeaderX, currentY, valueColW, rowHeight));
            requests.push(RequestFactory.updateShapeProperties(lvId, rowBg, theme.colors.faintGray, 1, 'MIDDLE'));
            requests.push(...BatchTextStyleUtils.setText(slideId, lvId, leftValue, { size: theme.fonts.sizes.body, align: 'CENTER', color: compareColors.left }, theme));

            // Right value cell
            const rvId = baseId + '_RV';
            requests.push(RequestFactory.createShape(slideId, rvId, 'RECTANGLE', rightHeaderX, currentY, valueColW - (trend ? layout.pxToPt(40) : 0), rowHeight));
            requests.push(RequestFactory.updateShapeProperties(rvId, rowBg, theme.colors.faintGray, 1, 'MIDDLE'));
            requests.push(...BatchTextStyleUtils.setText(slideId, rvId, rightValue, { size: theme.fonts.sizes.body, align: 'CENTER', color: compareColors.right }, theme));

            // Trend
            if (trend) {
                const trendX = rightHeaderX + valueColW - layout.pxToPt(35);
                const tId = baseId + '_TR';
                const isUp = String(trend).toLowerCase() === 'up';
                const trendColor = isUp ? '#28a745' : '#dc3545';

                requests.push(RequestFactory.createShape(slideId, tId, 'ELLIPSE', trendX, currentY + rowHeight / 4, layout.pxToPt(25), layout.pxToPt(25)));
                requests.push(RequestFactory.updateShapeProperties(tId, trendColor, 'TRANSPARENT', 0, 'MIDDLE'));
                requests.push(...BatchTextStyleUtils.setText(slideId, tId, isUp ? '↑' : '↓', { size: 12, color: '#FFFFFF', bold: true, align: 'CENTER' }, theme));
            }

            currentY += rowHeight;
        });

        return requests;
    }
}
