import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class TriangleDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return requests;

        const itemsToDraw = items.slice(0, 3);

        const centerX = area.left + area.width / 2;
        const centerY = area.top + area.height / 2;
        const radius = Math.min(area.width, area.height) / 3.2;

        const positions = [
            { x: centerX, y: centerY - radius },
            { x: centerX + radius * 0.866, y: centerY + radius * 0.5 },
            { x: centerX - radius * 0.866, y: centerY + radius * 0.5 }
        ];

        const circleSize = layout.pxToPt(160);

        // Center Triangle (Shape)
        const triangleId = slideId + '_TRI_CENTER';
        requests.push(RequestFactory.createShape(slideId, triangleId, 'TRIANGLE', centerX - radius, centerY - radius, radius * 2, radius * 2));
        requests.push(RequestFactory.updateShapeProperties(triangleId, theme.colors.faintGray || '#EEEEEE', 'TRANSPARENT'));
        // Note: Rotation is tricky in Batch API createShape. Default is 0.
        // Assuming TRIANGLE shape points UP by default. 
        // We positioned it bounding box center.

        itemsToDraw.forEach((item: any, i: number) => {
            const pos = positions[i];
            const title = item.title || item.label || '';
            const desc = item.desc || item.subLabel || '';

            const x = pos.x - circleSize / 2;
            const y = pos.y - circleSize / 2;

            const circleId = slideId + '_TRI_ITEM_' + i;

            // Circle Shape
            requests.push(RequestFactory.createShape(slideId, circleId, 'ELLIPSE', x, y, circleSize, circleSize));
            requests.push(RequestFactory.updateShapeProperties(circleId, settings.primaryColor, 'TRANSPARENT', 0, 'MIDDLE'));

            // Text
            const fullText = `${title}\n${desc}`;
            requests.push(RequestFactory.insertText(circleId, fullText));

            // Style Title
            if (title.length > 0) {
                requests.push(RequestFactory.updateTextStyle(circleId, {
                    fontSize: 14,
                    bold: true,
                    color: theme.colors.backgroundGray,
                    fontFamily: theme.fonts.family
                }, 0, title.length));
            }
            // Style Desc
            if (desc.length > 0) {
                requests.push(RequestFactory.updateTextStyle(circleId, {
                    fontSize: 12,
                    bold: false,
                    color: theme.colors.backgroundGray,
                    fontFamily: theme.fonts.family
                }, title.length + 1, desc.length));
            }

            requests.push(RequestFactory.updateParagraphStyle(circleId, 'CENTER'));
        });

        return requests;
    }
}
