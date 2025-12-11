import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class CycleDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return requests;

        const centerX = area.left + area.width / 2;
        const centerY = area.top + area.height / 2;
        const orbitRadius = Math.min(area.width, area.height) * 0.35;
        const count = items.length;
        const angleStep = (2 * Math.PI) / count;
        const startAngle = -Math.PI / 2;

        // 1. Connectors (Thick Spoke lines)
        items.forEach((item: any, i: number) => {
            const angle = startAngle + (i * angleStep);
            const x = centerX + Math.cos(angle) * orbitRadius;
            const y = centerY + Math.sin(angle) * orbitRadius;

            const connId = slideId + `_CYC_CONN_${i}`;
            // Draw line from Center border to Node border?
            // Simple approach: Center to Node center, behind everything.
            requests.push(RequestFactory.createLine(slideId, connId, centerX, centerY, x, y));
            requests.push({
                updateLineProperties: {
                    objectId: connId,
                    lineProperties: {
                        lineFill: { solidFill: { color: { rgbColor: { red: 0.8, green: 0.8, blue: 0.8 } } } }, // Light Gray
                        weight: { magnitude: 4, unit: 'PT' }, // Thicker
                        dashStyle: 'SOLID'
                    },
                    fields: 'lineFill,weight,dashStyle'
                }
            });
        });

        // 2. Center Hub (Premium Style)
        const centerR = 140; // Larger
        const centerId = slideId + '_CYC_CENTER';

        // Shadow effect (Gray circle slightly offset)
        const shadowId = slideId + '_CYC_CENTER_SHADOW';
        requests.push(RequestFactory.createShape(slideId, shadowId, 'ELLIPSE', centerX - centerR / 2 + 3, centerY - centerR / 2 + 3, centerR, centerR));
        requests.push(RequestFactory.updateShapeProperties(shadowId, '#DDDDDD', 'TRANSPARENT'));

        // Main Circle
        requests.push(RequestFactory.createShape(slideId, centerId, 'ELLIPSE', centerX - centerR / 2, centerY - centerR / 2, centerR, centerR));
        requests.push(RequestFactory.updateShapeProperties(centerId, settings.primaryColor, '#FFFFFF', 3)); // White border for pop

        const centerText = data.centerText || data.title || 'CYCLE';
        requests.push({ insertText: { objectId: centerId, text: centerText } });
        requests.push(RequestFactory.updateTextStyle(centerId, {
            fontSize: 20,
            bold: true,
            color: '#FFFFFF',
            fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: centerId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
        requests.push({ updateShapeProperties: { objectId: centerId, shapeProperties: { contentAlignment: 'MIDDLE' as any }, fields: 'contentAlignment' } });

        // 3. Satellite Nodes
        items.forEach((item: any, i: number) => {
            const angle = startAngle + (i * angleStep);
            const x = centerX + Math.cos(angle) * orbitRadius;
            const y = centerY + Math.sin(angle) * orbitRadius;

            const unique = `_CYC_${i}`;
            const bubbleId = slideId + unique + '_BUBBLE';
            const labelId = slideId + unique + '_LABEL';

            // Bubble (Solid Color now, not outline)
            const bubbleR = 60;
            // Shadow
            const bShadowId = bubbleId + '_SHADOW';
            requests.push(RequestFactory.createShape(slideId, bShadowId, 'ELLIPSE', x - bubbleR / 2 + 2, y - bubbleR / 2 + 2, bubbleR, bubbleR));
            requests.push(RequestFactory.updateShapeProperties(bShadowId, '#DDDDDD', 'TRANSPARENT'));

            // Main Bubble
            requests.push(RequestFactory.createShape(slideId, bubbleId, 'ELLIPSE', x - bubbleR / 2, y - bubbleR / 2, bubbleR, bubbleR));
            requests.push(RequestFactory.updateShapeProperties(bubbleId, '#FFFFFF', settings.primaryColor, 2)); // White with colored border

            const numStr = String(i + 1).padStart(2, '0');
            requests.push({ insertText: { objectId: bubbleId, text: numStr } });
            requests.push(RequestFactory.updateTextStyle(bubbleId, {
                fontSize: 24,
                bold: true,
                color: settings.primaryColor,
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: bubbleId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: bubbleId, shapeProperties: { contentAlignment: 'MIDDLE' as any }, fields: 'contentAlignment' } });

            // Label
            const isRight = x > centerX;
            const textW = 140;
            const textH = 70;
            // Adjust position to avoid overlapping bubble
            const offsetX = isRight ? bubbleR / 2 + 5 : -bubbleR / 2 - textW - 5;
            const textX = x + offsetX;
            const textY = y - textH / 2;

            const label = (typeof item === 'string') ? item : (item.label || item.title || '');
            requests.push(RequestFactory.createShape(slideId, labelId, 'TEXT_BOX', textX, textY, textW, textH));
            requests.push({ insertText: { objectId: labelId, text: label } });
            requests.push(RequestFactory.updateTextStyle(labelId, {
                fontSize: 14,
                bold: true,
                color: theme.colors.textPrimary || '#333333',
                fontFamily: theme.fonts.family
            }));
            requests.push({
                updateParagraphStyle: {
                    objectId: labelId,
                    style: { alignment: isRight ? 'START' : 'END' },
                    fields: 'alignment'
                }
            });
            requests.push({ updateShapeProperties: { objectId: labelId, shapeProperties: { contentAlignment: 'MIDDLE' as any }, fields: 'contentAlignment' } });
        });

        return requests;
    }
}
