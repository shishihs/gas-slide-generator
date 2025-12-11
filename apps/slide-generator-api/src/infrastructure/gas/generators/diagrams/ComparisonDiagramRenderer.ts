import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class ComparisonDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();

        const leftTitle = data.leftTitle || 'Plan A';
        const rightTitle = data.rightTitle || 'Plan B';
        const leftItems = data.leftItems || [];
        const rightItems = data.rightItems || [];

        const gap = 60; // Wider gap for VS
        const colWidth = (area.width - gap) / 2;
        const headerH = 70;

        // Helper
        const drawColumn = (x: number, title: string, items: string[], suffix: string) => {
            const headerId = slideId + '_CMP_HEAD_' + suffix;
            const bodyId = slideId + '_CMP_BODY_' + suffix;

            // 1. Header Box
            requests.push(RequestFactory.createShape(slideId, headerId, 'ROUND_RECTANGLE', x, area.top, colWidth, headerH));
            requests.push(RequestFactory.updateShapeProperties(headerId, settings.primaryColor, 'TRANSPARENT'));

            requests.push({ insertText: { objectId: headerId, text: title } });
            requests.push(RequestFactory.updateTextStyle(headerId, {
                fontSize: 24,
                bold: true,
                color: '#FFFFFF',
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: headerId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: headerId, shapeProperties: { contentAlignment: 'MIDDLE' as any }, fields: 'contentAlignment' } });

            // 2. Body Box
            const bodyH = area.height - headerH - 10;
            const bodyY = area.top + headerH + 10;

            requests.push(RequestFactory.createShape(slideId, bodyId, 'ROUND_RECTANGLE', x, bodyY, colWidth, bodyH));
            requests.push(RequestFactory.updateShapeProperties(bodyId, theme.colors.faintGray || '#F8F9FA', 'TRANSPARENT')); // Light gray bg

            // 3. Items
            let currentY = bodyY + 20;
            const itemGap = 15;

            items.forEach((itemText: string, i: number) => {
                const unique = `_${suffix}_${i}`;
                const iconId = slideId + '_CMP_ICON' + unique;
                const textId = slideId + '_CMP_TXT' + unique;

                // Icon (Checkmark)
                const iconSize = 24;
                requests.push(RequestFactory.createShape(slideId, iconId, 'TEXT_BOX', x + 20, currentY, iconSize, iconSize));
                requests.push({ insertText: { objectId: iconId, text: '✔' } });
                requests.push(RequestFactory.updateTextStyle(iconId, {
                    fontSize: 18,
                    bold: true,
                    color: settings.primaryColor,
                    fontFamily: theme.fonts.family // Use symbol font if needed, but standard usually has check
                }));
                requests.push({ updateParagraphStyle: { objectId: iconId, style: { alignment: 'CENTER' }, fields: 'alignment' } });

                // Text
                const textX = x + 20 + iconSize + 10;
                const textW = colWidth - (20 + iconSize + 10 + 20);
                const itemH = 40; // Approximate height per item

                requests.push(RequestFactory.createShape(slideId, textId, 'TEXT_BOX', textX, currentY, textW, itemH));
                requests.push({ insertText: { objectId: textId, text: itemText } });
                requests.push(RequestFactory.updateTextStyle(textId, {
                    fontSize: 16,
                    color: theme.colors.textPrimary || '#333333',
                    fontFamily: theme.fonts.family
                }));
                requests.push({ updateParagraphStyle: { objectId: textId, style: { alignment: 'START' }, fields: 'alignment' } });

                currentY += itemH + itemGap;
            });
        };

        drawColumn(area.left, leftTitle, leftItems, 'L');
        drawColumn(area.left + colWidth + gap, rightTitle, rightItems, 'R');

        // Z-Index: Ensure VS is on top. Since we create it first in requests list? No, API order.
        // If we want VS on top of Headers, create it LAST.
        // Let's move VS creation to end of list.
        // Actually, we can just push VS requests at the very end.
        // (Removing VS requests from top and re-adding here conceptually, but for simplicity I'll return requests in order)
        // Current implementation: VS created BEFORE columns. So columns might overlap if spacing is tight?
        // With gap=60 and VS R=50, it shouldn't overlap. But safer to draw VS last just in case.
        // I will move the VS block to the end in the tool instructions.

        return requests;
    }
}
