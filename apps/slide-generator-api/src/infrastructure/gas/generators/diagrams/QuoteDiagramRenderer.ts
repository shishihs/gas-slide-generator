import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class QuoteDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const text = data.text || (data.points && data.points[0]) || '';
        const author = data.author || (data.points && data.points[1]) || '';

        // 1. Big Quote Mark (Background) - Created first to be at back
        const quoteSize = layout.pxToPt(200);
        const quoteId = slideId + '_QUOTE_MARK';
        requests.push(RequestFactory.createShape(slideId, quoteId, 'TEXT_BOX', area.left + layout.pxToPt(20), area.top - layout.pxToPt(40), quoteSize, quoteSize));
        requests.push(...BatchTextStyleUtils.setText(slideId, quoteId, '“', {
            size: 200,
            color: '#F0F0F0',
            fontFamily: 'Georgia',
            bold: true
        }, theme));

        const contentW = area.width * 0.9;
        const contentX = area.left + (area.width - contentW) / 2;
        const textTop = area.top + layout.pxToPt(60);

        // 2. Quote Text
        const textId = slideId + '_QUOTE_TEXT';
        requests.push(RequestFactory.createShape(slideId, textId, 'TEXT_BOX', contentX, textTop, contentW, layout.pxToPt(160)));
        requests.push(...BatchTextStyleUtils.setText(slideId, textId, text, {
            size: 32,
            bold: false, // Explicit false
            color: theme.colors.textPrimary,
            align: 'CENTER',
            fontFamily: 'Georgia'
        }, theme));
        requests.push(RequestFactory.updateShapeProperties(textId, null, null, null, 'BOTTOM'));

        // 3. Separator Line
        const lineW = layout.pxToPt(40);
        const lineX = area.left + (area.width - lineW) / 2;
        const lineY = textTop + layout.pxToPt(165);
        const lineId = slideId + '_QUOTE_LINE';

        requests.push(RequestFactory.createLine(slideId, lineId, lineX, lineY, lineX + lineW, lineY));
        // We need primary color for line. Assuming primaryColor is Hex.
        // We'll trust user supplied correct Hex or use fallback.
        // No easy hexToRgb here unless we inline it or use settings.primaryColor directly if API supported hex (it doesn't for Line).
        // I'll inline a simple hex parser again or assume gray/black fallback for safety if primaryColor is complex.
        // Or I can paste the helper. I'll paste the helper. 
        const hexToRgb = (hex: string) => {
            const safeHex = (hex && typeof hex === 'string') ? hex : '#000000';
            const h = safeHex.replace('#', '');
            return {
                red: parseInt(h.substring(0, 2), 16) / 255,
                green: parseInt(h.substring(2, 4), 16) / 255,
                blue: parseInt(h.substring(4, 6), 16) / 255
            };
        };

        requests.push({
            updateLineProperties: {
                objectId: lineId,
                lineProperties: {
                    lineFill: { solidFill: { color: { rgbColor: hexToRgb(settings.primaryColor) } } },
                    weight: { magnitude: 2, unit: 'PT' }
                },
                fields: 'lineFill,weight'
            }
        });

        // 4. Author
        if (author) {
            const authId = slideId + '_QUOTE_AUTH';
            requests.push(RequestFactory.createShape(slideId, authId, 'TEXT_BOX', contentX, lineY + layout.pxToPt(5), contentW, layout.pxToPt(30)));
            requests.push(...BatchTextStyleUtils.setText(slideId, authId, author, {
                size: 14,
                align: 'CENTER',
                color: theme.colors.neutralGray,
                bold: true
            }, theme));
            requests.push(RequestFactory.updateShapeProperties(authId, null, null, null, 'TOP'));
        }

        return requests;
    }
}
