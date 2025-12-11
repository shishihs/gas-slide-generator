
import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class TimelineDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const milestones = data.milestones || data.items || [];

        if (milestones.length === 0) return requests;

        const centerY = area.top + area.height / 2;
        const startX = area.left + 50;
        const endX = area.left + area.width - 50;

        // 1. Central Axis
        const axisId = slideId + '_TL_AXIS';
        requests.push(RequestFactory.createLine(slideId, axisId, startX, centerY, endX, centerY));
        // Style Axis
        requests.push({
            updateLineProperties: {
                objectId: axisId,
                lineProperties: {
                    weight: { magnitude: 4, unit: 'PT' },
                    lineFill: { solidFill: { color: { rgbColor: { red: 0.8, green: 0.8, blue: 0.8 } } } }, // neutralGray #CCCCCC
                    dashStyle: 'SOLID'
                },
                fields: 'weight,lineFill,dashStyle'
            }
        });

        // 2. Milestones
        const count = milestones.length;
        const usableWidth = endX - startX;
        const gap = usableWidth / (count + 1);

        milestones.forEach((m: any, i: number) => {
            const x = startX + (gap * (i + 1));
            const isTop = (i % 2 === 0);
            const cardH = 80;
            const cardW = 140;
            const connLen = 50; // slightly longer
            const dotR = 24; // larger dot

            const unique = `_${i}`;
            const dotId = slideId + '_TL_DOT' + unique;
            const connId = slideId + '_TL_CONN' + unique;
            const cardId = slideId + '_TL_CARD' + unique;

            // Connector Y
            const connY1 = centerY;
            const connY2 = isTop ? centerY - connLen : centerY + connLen;

            // Dot (Central Node)
            requests.push(RequestFactory.createShape(slideId, dotId, 'ELLIPSE', x - dotR / 2, centerY - dotR / 2, dotR, dotR));
            // Style Dot
            requests.push(RequestFactory.updateShapeProperties(dotId, settings.primaryColor, '#FFFFFF', 3)); // White border for contrast

            // Connector Line
            requests.push(RequestFactory.createLine(slideId, connId, x, connY1, x, connY2));
            requests.push({
                updateLineProperties: {
                    objectId: connId,
                    lineProperties: {
                        weight: { magnitude: 2, unit: 'PT' },
                        lineFill: { solidFill: { color: { rgbColor: { red: 0.7, green: 0.7, blue: 0.7 } } } },
                        dashStyle: 'DOT'
                    },
                    fields: 'weight,lineFill,dashStyle'
                }
            });

            // Card
            const cardY = isTop ? connY2 - cardH : connY2;
            const cardX = x - cardW / 2;

            requests.push(RequestFactory.createShape(slideId, cardId, 'ROUND_RECTANGLE', cardX, cardY, cardW, cardH));
            // Card Style
            requests.push(RequestFactory.updateShapeProperties(cardId, '#FFFFFF', theme.colors.ghostGray));

            // Content
            const date = m.date || m.year || '';
            const label = m.label || m.title || m.text || '';
            const fullText = `${date}\n${label}`;

            requests.push({ insertText: { objectId: cardId, text: fullText } });

            // Style Date (bold, colored)
            if (date) {
                // Use RequestFactory helper
                requests.push(RequestFactory.updateTextStyle(cardId, {
                    bold: true,
                    fontSize: 14,
                    color: settings.primaryColor,
                    fontFamily: theme.fonts.family
                }, 0, date.length));
            }

            // Style Label (normal)
            if (label) {
                const start = date.length + 1; // +1 for newline
                requests.push(RequestFactory.updateTextStyle(cardId, {
                    bold: false,
                    fontSize: 12,
                    color: theme.colors.textPrimary || '#333333',
                    fontFamily: theme.fonts.family
                }, start, label.length));
            }

            // Alignment Center
            requests.push({ updateParagraphStyle: { objectId: cardId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: cardId, shapeProperties: { contentAlignment: 'MIDDLE' as any }, fields: 'contentAlignment' } });
        });

        return requests;
    }
}
