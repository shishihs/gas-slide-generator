import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class AgendaDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items: string[] = data.items || [];
        if (items.length === 0) return requests;

        // Configuration
        const COLUMN_COUNT = 2;
        const ROW_COUNT = Math.ceil(items.length / COLUMN_COUNT);

        // Use full width of the area, but add padding between items
        const GAP_X = 40;
        const GAP_Y = 25;

        const itemWidth = (area.width - (GAP_X * (COLUMN_COUNT - 1))) / COLUMN_COUNT;
        const itemHeight = Math.min((area.height - (GAP_Y * (ROW_COUNT - 1))) / ROW_COUNT, 80);

        items.forEach((itemText: string, index: number) => {
            const col = index % COLUMN_COUNT;
            const row = Math.floor(index / COLUMN_COUNT);

            const x = area.left + (col * (itemWidth + GAP_X));
            const y = area.top + (row * (itemHeight + GAP_Y));

            const uniqueId = `_AG_${index}`;
            const numShapeId = slideId + uniqueId + '_NUM';
            const cardShapeId = slideId + uniqueId + '_CARD';

            // 2. Number Box (Square on the left)
            const numberSize = itemHeight; // Square
            requests.push(RequestFactory.createShape(slideId, numShapeId, 'RECTANGLE', x, y, numberSize, numberSize));
            requests.push(RequestFactory.updateShapeProperties(numShapeId, settings.primaryColor, 'TRANSPARENT'));

            const numStr = (index + 1).toString().padStart(2, '0');
            requests.push({ insertText: { objectId: numShapeId, text: numStr } });
            requests.push(RequestFactory.updateTextStyle(numShapeId, {
                fontSize: 28,
                bold: true,
                color: '#FFFFFF',
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: numShapeId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: numShapeId, shapeProperties: { contentAlignment: 'MIDDLE' as any }, fields: 'contentAlignment' } });

            // 3. Text Box (Right side)
            const textWidth = itemWidth - numberSize;
            const textX = x + numberSize;

            // Background for text (Subtle card)
            requests.push(RequestFactory.createShape(slideId, cardShapeId, 'RECTANGLE', textX, y, textWidth, itemHeight));
            requests.push(RequestFactory.updateShapeProperties(cardShapeId, settings.card_bg || '#F5F5F5', 'TRANSPARENT'));

            // Text content
            requests.push({ insertText: { objectId: cardShapeId, text: itemText } });
            requests.push(RequestFactory.updateTextStyle(cardShapeId, {
                fontSize: 18,
                color: settings.text_primary || '#333333',
                fontFamily: theme.fonts.family
            }));
            requests.push({
                updateShapeProperties: {
                    objectId: cardShapeId,
                    shapeProperties: { contentAlignment: 'MIDDLE' as any },
                    fields: 'contentAlignment'
                }
            });
        });

        return requests;
    }
}
