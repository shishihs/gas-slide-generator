import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class FAQDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items = data.items || data.points || [];

        const parsedItems: any[] = [];
        if (items.length && typeof items[0] === 'string') {
            let currentQ = '';
            items.forEach((str: string) => {
                if (str.startsWith('Q:') || str.startsWith('Q.')) currentQ = str;
                else if (str.startsWith('A:') || str.startsWith('A.')) parsedItems.push({ q: currentQ, a: str });
            });
        } else {
            items.forEach((it: any) => parsedItems.push(it));
        }

        if (!parsedItems.length) return requests;

        const gap = layout.pxToPt(30);
        const itemH = (area.height - (gap * (parsedItems.length - 1))) / parsedItems.length;

        parsedItems.forEach((item, i) => {
            const y = area.top + i * (itemH + gap);
            const qStr = (item.q || '').replace(/^[QA][:. ]+/, '');
            const aStr = (item.a || '').replace(/^[QA][:. ]+/, '');
            const baseId = slideId + `_FAQ_${i}`;

            // Q Indicator
            const qIndId = baseId + '_QI';
            requests.push(RequestFactory.createShape(slideId, qIndId, 'TEXT_BOX', area.left, y, layout.pxToPt(30), layout.pxToPt(30)));
            requests.push(...BatchTextStyleUtils.setText(slideId, qIndId, 'Q.', { size: 16, bold: true, color: settings.primaryColor }, theme));

            // Q Content
            const qBoxId = baseId + '_QC';
            requests.push(RequestFactory.createShape(slideId, qBoxId, 'TEXT_BOX', area.left + layout.pxToPt(30), y, area.width - layout.pxToPt(30), layout.pxToPt(40)));
            requests.push(...BatchTextStyleUtils.setText(slideId, qBoxId, qStr, { size: 14, bold: true, color: theme.colors.textPrimary }, theme));

            // A Indicator
            const aY = y + layout.pxToPt(30);
            const aIndId = baseId + '_AI';
            requests.push(RequestFactory.createShape(slideId, aIndId, 'TEXT_BOX', area.left, aY, layout.pxToPt(30), layout.pxToPt(30)));
            requests.push(...BatchTextStyleUtils.setText(slideId, aIndId, 'A.', { size: 16, bold: true, color: theme.colors.neutralGray }, theme));

            // A Content
            const aBoxHasHeight = itemH - layout.pxToPt(40);
            if (aBoxHasHeight > 10) {
                const aBoxId = baseId + '_AC';
                requests.push(RequestFactory.createShape(slideId, aBoxId, 'TEXT_BOX', area.left + layout.pxToPt(30), aY, area.width - layout.pxToPt(30), aBoxHasHeight));
                requests.push(...BatchTextStyleUtils.setText(slideId, aBoxId, aStr, { size: 12, color: theme.colors.textPrimary }, theme));
                requests.push(RequestFactory.updateShapeProperties(aBoxId, null, null, null, 'TOP'));
            }

            // Separator Line
            if (i < parsedItems.length - 1) {
                const lineY = y + itemH + gap / 2;
                const lineId = baseId + '_LN';
                requests.push(RequestFactory.createLine(slideId, lineId, area.left, lineY, area.left + area.width, lineY));
                requests.push({
                    updateLineProperties: {
                        objectId: lineId,
                        lineProperties: {
                            lineFill: { solidFill: { color: { rgbColor: { red: 0.9, green: 0.9, blue: 0.9 } } } }, // Approx Ghost Gray
                            weight: { magnitude: 0.5, unit: 'PT' }
                        },
                        fields: 'lineFill,weight'
                    }
                });
            }
        });

        return requests;
    }
}
