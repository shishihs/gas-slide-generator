import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class ProgressDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return requests;

        const rowH = layout.pxToPt(50);
        const gap = layout.pxToPt(15);
        const startY = area.top + (area.height - (items.length * (rowH + gap))) / 2;

        items.forEach((item: any, i: number) => {
            const y = startY + i * (rowH + gap);
            const labelW = layout.pxToPt(150);
            const barAreaW = area.width - labelW - layout.pxToPt(60);
            const baseId = slideId + `_PROG_${i}`;

            // Label
            const labelId = baseId + '_LBL';
            requests.push(RequestFactory.createShape(slideId, labelId, 'TEXT_BOX', area.left, y, labelW, rowH));
            requests.push(...BatchTextStyleUtils.setText(slideId, labelId, item.label || '', {
                size: 14, bold: true, align: 'END'
            }, theme));
            requests.push(RequestFactory.updateShapeProperties(labelId, null, null, null, 'MIDDLE'));

            // Bar BG
            const barBgId = baseId + '_BG';
            requests.push(RequestFactory.createShape(slideId, barBgId, 'ROUND_RECTANGLE', area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW, rowH / 3));
            requests.push(RequestFactory.updateShapeProperties(barBgId, theme.colors.ghostGray, 'TRANSPARENT'));

            // Bar FG
            const percent = Math.min(100, Math.max(0, parseInt(item.percent || 0)));
            if (percent > 0) {
                const barFgId = baseId + '_FG';
                const fgW = barAreaW * (percent / 100);
                // Ensure width is at least something if percent > 0? No, percent is enough.
                requests.push(RequestFactory.createShape(slideId, barFgId, 'ROUND_RECTANGLE', area.left + labelW + layout.pxToPt(20), y + rowH / 3, fgW, rowH / 3));
                requests.push(RequestFactory.updateShapeProperties(barFgId, settings.primaryColor, 'TRANSPARENT'));
            }

            // Value
            const valId = baseId + '_VAL';
            requests.push(RequestFactory.createShape(slideId, valId, 'TEXT_BOX', area.left + labelW + barAreaW + layout.pxToPt(30), y, layout.pxToPt(50), rowH));
            requests.push(...BatchTextStyleUtils.setText(slideId, valId, `${percent}%`, {
                size: 14, color: theme.colors.neutralGray
            }, theme));
            requests.push(RequestFactory.updateShapeProperties(valId, null, null, null, 'MIDDLE'));
        });

        return requests;
    }
}
