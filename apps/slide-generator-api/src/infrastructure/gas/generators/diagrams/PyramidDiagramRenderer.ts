import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class PyramidDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const levels = data.levels || data.items || [];
        if (!levels.length) return requests;

        const count = levels.length;
        const pyramidH = Math.min(area.height, 400);
        const pyramidW = Math.min(area.width * 0.6, 500);

        const centerX = area.left + (area.width / 2);
        const topY = area.top + (area.height - pyramidH) / 2;

        // 1. Main Triangle
        const triId = slideId + '_PYR_MAIN';
        requests.push(RequestFactory.createShape(slideId, triId, 'TRIANGLE', centerX - pyramidW / 2, topY, pyramidW, pyramidH));
        requests.push(RequestFactory.updateShapeProperties(triId, settings.primaryColor, 'TRANSPARENT'));

        // 2. White Lines for slicing
        // Gap between levels
        const gap = 4;
        const levelHeight = (pyramidH - (gap * (count - 1))) / count;

        for (let i = 1; i < count; i++) {
            const lineY = topY + (i * (pyramidH / count));
            const ratio = i / count;
            const widthAtY = pyramidW * ratio;

            const lineId = slideId + `_PYR_LINE_${i}`;
            const lineW = widthAtY - 4; // slight indent
            const lineH = 2; // Thickness
            const lineLeft = centerX - lineW / 2;

            // Use Rectangle as a "thick line" eraser
            requests.push(RequestFactory.createShape(slideId, lineId, 'RECTANGLE', lineLeft, lineY, lineW, lineH));
            requests.push(RequestFactory.updateShapeProperties(lineId, '#FFFFFF', 'TRANSPARENT'));
        }

        // 3. Labels
        levels.forEach((level: any, index: number) => {
            const y = topY + (index * (pyramidH / count)) + (levelHeight / 2) - 10;

            // Calc X on right edge
            // Linear width calc: w at y = (y - top) / H * W? No, triangle widens.
            // Width at depth d (0..1) = d * W? No, apex is top.
            // Width at depth d (from top) = d * W.
            const depthRatio = (index + 0.5) / count;
            const widthAtDepth = pyramidW * depthRatio;
            const pyramidEdgeX = centerX + widthAtDepth / 2;

            const lineStartX = pyramidEdgeX + 10;
            const lineEndX = lineStartX + 30;
            const textX = lineEndX + 10;

            // Connector
            const connId = slideId + `_PYR_CONN_${index}`;
            requests.push(RequestFactory.createLine(slideId, connId, lineStartX, y + 10, lineEndX, y + 10));
            requests.push({
                updateLineProperties: {
                    objectId: connId,
                    lineProperties: {
                        lineFill: { solidFill: { color: { rgbColor: { red: 0.2, green: 0.2, blue: 0.2 } } } }, // Dark Gray
                        weight: { magnitude: 1, unit: 'PT' }
                    },
                    fields: 'lineFill,weight'
                }
            });

            // Text Box
            const textId = slideId + `_PYR_TXT_${index}`;
            const title = level.label || level.title || '';
            const sub = level.subLabel || level.desc || '';
            const fullText = `${title}\n${sub}`;

            requests.push(RequestFactory.createShape(slideId, textId, 'TEXT_BOX', textX, y - 25, 200, 60)); // Adjust Y centering
            requests.push({ insertText: { objectId: textId, text: fullText } });

            // Style Title
            if (title) {
                requests.push(RequestFactory.updateTextStyle(textId, {
                    bold: true,
                    fontSize: 16,
                    color: settings.text_primary || '#333333',
                    fontFamily: layout.getTheme().fonts.family
                }, 0, title.length));
            }

            // Style Subtitle
            if (sub) {
                const subStart = title.length + 1;
                requests.push(RequestFactory.updateTextStyle(textId, {
                    bold: false,
                    fontSize: 12,
                    color: settings.ghost_gray || '#666666',
                    fontFamily: layout.getTheme().fonts.family
                }, subStart, sub.length));
            }

            // Vertical Align Middle
            requests.push({
                updateShapeProperties: {
                    objectId: textId,
                    shapeProperties: { contentAlignment: 'MIDDLE' as any },
                    fields: 'contentAlignment'
                }
            });
        });

        return requests;
    }
}
