import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class StepUpDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return requests;
        const count = items.length;
        const validCount = Math.min(count, 5);

        const gap = 20;
        const barWidth = (area.width - (gap * (validCount - 1))) / validCount;

        // Max height for the tallest bar
        const maxBarH = area.height * 0.7;
        const minBarH = area.height * 0.3; // slightly taller min

        items.slice(0, validCount).forEach((item: any, i: number) => {
            // Linear growth
            const ratio = i / (validCount - 1 || 1);
            const barH = minBarH + (maxBarH - minBarH) * ratio;

            const x = area.left + (i * (barWidth + gap));
            const y = area.top + area.height - barH;

            const unique = `_STEP_${i}`;
            const barId = slideId + unique + '_BAR';
            const numId = slideId + unique + '_NUM';
            const titleId = slideId + unique + '_TITLE';
            const descId = slideId + unique + '_DESC';

            // 1. Bar Shape
            requests.push(RequestFactory.createShape(slideId, barId, 'RECTANGLE', x, y, barWidth, barH));
            // Color Gradient or Solid? Solid primary for now, maybe darken as steps go up?
            // "Step Up" often implies intensity.
            requests.push(RequestFactory.updateShapeProperties(barId, settings.primaryColor, 'TRANSPARENT'));

            // 2. Number (Bottom of Bar, large and subtle)
            const numStr = String(i + 1).padStart(2, '0');
            requests.push(RequestFactory.createShape(slideId, numId, 'TEXT_BOX', x, y + barH - 50, barWidth, 40));
            requests.push({ insertText: { objectId: numId, text: numStr } });
            requests.push(RequestFactory.updateTextStyle(numId, {
                fontSize: 32,
                bold: true,
                color: '#FFFFFF', // High contrast
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: numId, style: { alignment: 'CENTER' }, fields: 'alignment' } });

            // Opacity for number? API doesn't do text opacity easily locally. Keep solid white.

            // 3. Title (Above Bar - Important)
            // If Text is object, extract .title or .text
            const titleContent = (typeof item === 'string') ? item : (item.title || item.label || '');

            requests.push(RequestFactory.createShape(slideId, titleId, 'TEXT_BOX', x, y - 60, barWidth, 50));
            requests.push({ insertText: { objectId: titleId, text: titleContent } });
            requests.push(RequestFactory.updateTextStyle(titleId, {
                fontSize: 16,
                bold: true,
                color: theme.colors.textPrimary || '#333333',
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: 'BOTTOM' as any }, fields: 'contentAlignment' } });

            // 4. Description (Inside Bar, Top)
            // Only if space permits
            const desc = (typeof item === 'object') ? (item.desc || item.description) : null;
            if (desc) {
                const descBoxH = Math.min(barH - 60, 100);
                if (descBoxH > 30) {
                    requests.push(RequestFactory.createShape(slideId, descId, 'TEXT_BOX', x + 5, y + 10, barWidth - 10, descBoxH));
                    requests.push({ insertText: { objectId: descId, text: desc } });
                    requests.push(RequestFactory.updateTextStyle(descId, {
                        fontSize: 12,
                        color: '#FFFFFF',
                        fontFamily: theme.fonts.family
                    }));
                    requests.push({ updateParagraphStyle: { objectId: descId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
                    requests.push({ updateShapeProperties: { objectId: descId, shapeProperties: { contentAlignment: 'TOP' as any }, fields: 'contentAlignment' } });
                }
            }
        });

        return requests;
    }
}
