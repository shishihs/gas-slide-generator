import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class CardsDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return requests;
        const type = (data.type || '').toLowerCase();
        const hasHeader = type.includes('headercards');

        const cols = data.columns || Math.min(items.length, 3);
        const rows = Math.ceil(items.length / cols);

        const gap = layout.pxToPt(30);
        const cardW = (area.width - gap * (cols - 1)) / cols;
        const cardH = (area.height - gap * (rows - 1)) / rows;

        items.forEach((item: any, i: number) => {
            const r = Math.floor(i / cols);
            const c = i % cols;
            const x = area.left + c * (cardW + gap);
            const y = area.top + r * (cardH + gap);

            let title = '';
            let desc = '';
            if (typeof item === 'string') {
                const lines = item.split('\n');
                title = lines[0] || '';
                desc = lines.slice(1).join('\n') || '';
            } else {
                title = item.title || item.label || '';
                desc = item.desc || item.description || item.text || '';
            }

            const baseId = slideId + '_CARD_' + i;

            if (hasHeader) {
                // Header Bar
                const barH = layout.pxToPt(4);
                const barId = baseId + '_BAR';
                requests.push(RequestFactory.createShape(slideId, barId, 'RECTANGLE', x, y, cardW, barH));
                requests.push(RequestFactory.updateShapeProperties(barId, settings.primaryColor, 'TRANSPARENT'));

                // Numbering
                const numStr = String(i + 1).padStart(2, '0');
                const numId = baseId + '_NUM';
                requests.push(RequestFactory.createShape(slideId, numId, 'TEXT_BOX', x, y + layout.pxToPt(6), cardW, layout.pxToPt(20)));
                requests.push(...BatchTextStyleUtils.setText(slideId, numId, numStr, {
                    size: 14, bold: true, color: theme.colors.neutralGray, align: 'END'
                }, theme));

                // Title
                const titleTop = y + layout.pxToPt(6);
                const titleH = layout.pxToPt(30);
                const titleId = baseId + '_TITLE';
                requests.push(RequestFactory.createShape(slideId, titleId, 'TEXT_BOX', x, titleTop, cardW, titleH));
                requests.push(RequestFactory.updateShapeProperties(titleId, null, null, null, 'TOP'));
                requests.push(...BatchTextStyleUtils.setText(slideId, titleId, title, {
                    size: 18, bold: true, color: theme.colors.textPrimary, align: 'START'
                }, theme));

                // Body
                const descTop = titleTop + titleH;
                const descH = cardH - (descTop - y);
                if (descH > 20) {
                    const descId = baseId + '_DESC';
                    requests.push(RequestFactory.createShape(slideId, descId, 'TEXT_BOX', x, descTop, cardW, descH));
                    requests.push(RequestFactory.updateShapeProperties(descId, null, null, null, 'TOP'));
                    requests.push(...BatchTextStyleUtils.setText(slideId, descId, desc, {
                        size: 13,
                        color: typeof theme.colors.textSmallFont === 'string' ? theme.colors.textSmallFont : '#424242',
                        align: 'START'
                    }, theme));
                }

            } else {
                // Minimal
                const dotSize = layout.pxToPt(6);
                const dotId = baseId + '_DOT';
                const dotY = y + layout.pxToPt(8);
                requests.push(RequestFactory.createShape(slideId, dotId, 'ELLIPSE', x, dotY, dotSize, dotSize));
                requests.push(RequestFactory.updateShapeProperties(dotId, settings.primaryColor, 'TRANSPARENT'));

                const contentX = x + dotSize + layout.pxToPt(10);
                const contentW = cardW - (dotSize + layout.pxToPt(10));

                const titleH = layout.pxToPt(30);
                const titleId = baseId + '_TITLE';
                requests.push(RequestFactory.createShape(slideId, titleId, 'TEXT_BOX', contentX, y, contentW, titleH));
                requests.push(RequestFactory.updateShapeProperties(titleId, null, null, null, 'TOP'));
                requests.push(...BatchTextStyleUtils.setText(slideId, titleId, title, {
                    size: 16, bold: true, color: theme.colors.textPrimary, align: 'START'
                }, theme));

                const descTop = y + titleH - layout.pxToPt(5);
                const descH = cardH - (descTop - y);
                if (descH > 20) {
                    const descId = baseId + '_DESC';
                    requests.push(RequestFactory.createShape(slideId, descId, 'TEXT_BOX', contentX, descTop, contentW, descH));
                    requests.push(RequestFactory.updateShapeProperties(descId, null, null, null, 'TOP'));
                    requests.push(...BatchTextStyleUtils.setText(slideId, descId, desc, {
                        size: 13,
                        color: typeof theme.colors.textSmallFont === 'string' ? theme.colors.textSmallFont : '#424242',
                        align: 'START'
                    }, theme));
                }
            }
        });

        return requests;
    }
}
