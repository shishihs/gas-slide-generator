import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class ProcessDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const steps = data.steps || data.items || [];
        if (!steps.length) return requests;

        const count = steps.length;
        const gap = -15; // Overlap for Chevron interlocking
        const itemWidth = (area.width - (gap * (count - 1))) / count;
        const itemHeight = Math.min(area.height, 150); // Limit height
        const y = area.top + (area.height - itemHeight) / 2; // Vertically center within area

        steps.forEach((step: string, i: number) => {
            const x = area.left + (i * (itemWidth + gap));
            const unique = `_${i}`;
            const shapeId = slideId + '_PR_CHEV_' + unique;

            // 1. Chevron Shape
            requests.push(RequestFactory.createShape(slideId, shapeId, 'CHEVRON', x, y, itemWidth, itemHeight));

            // Alternate colors
            const bgColor = i % 2 === 0 ? theme.colors.primary : theme.colors.deepPrimary || theme.colors.primary;

            requests.push({
                updateShapeProperties: {
                    objectId: shapeId,
                    shapeProperties: {
                        shapeBackgroundFill: { solidFill: { color: { rgbColor: RequestFactory.toRgbColor(bgColor) || { red: 0, green: 0, blue: 0 } } } },
                        outline: { propertyState: 'NOT_RENDERED' },
                        contentAlignment: 'MIDDLE'
                    },
                    fields: 'shapeBackgroundFill,outline,contentAlignment'
                }
            });

            // 2. Text Content
            // Step Number (Large) + Text
            // We can put them in the same text box (the shape itself)
            const cleanText = String(step || '').replace(/^\s*\d+[\.\s]*/, '');
            const fullText = `${i + 1}\n${cleanText}`;

            requests.push({ insertText: { objectId: shapeId, text: fullText } });

            // Style Number
            requests.push(RequestFactory.updateTextStyle(shapeId, {
                fontSize: 32,
                bold: true,
                color: '#FFFFFF',
                fontFamily: theme.fonts.family
            }, 0, String(i + 1).length));

            // Style Body
            requests.push(RequestFactory.updateTextStyle(shapeId, {
                fontSize: 14,
                bold: false,
                color: '#FFFFFF', // White text on colored background
                fontFamily: theme.fonts.family
            }, String(i + 1).length + 1, cleanText.length + 1)); // length fix? Call was (start, length)

            requests.push({
                updateParagraphStyle: {
                    objectId: shapeId,
                    style: { alignment: 'CENTER' },
                    fields: 'alignment'
                }
            });
        });

        return requests;
    }


}

