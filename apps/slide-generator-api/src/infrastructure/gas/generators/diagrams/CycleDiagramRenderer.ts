import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class CycleDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return;

        const centerX = area.left + area.width / 2;
        const centerY = area.top + area.height / 2;
        // Radius for the center of satellite items
        const orbitRadius = Math.min(area.width, area.height) * 0.35;

        // 1. Center Hub
        const centerR = 120;
        const centerBubble = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, centerX - centerR / 2, centerY - centerR / 2, centerR, centerR);
        centerBubble.getFill().setSolidFill(settings.primaryColor);
        centerBubble.getBorder().setTransparent();

        // Center Text
        const centerText = data.centerText || data.title || 'CYCLE';
        setStyledText(centerBubble, centerText, {
            size: 18,
            bold: true,
            color: '#FFFFFF',
            align: SlidesApp.ParagraphAlignment.CENTER
        }, theme);
        try { centerBubble.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        // 2. Satellite Items
        const count = items.length;
        const angleStep = (2 * Math.PI) / count;
        const startAngle = -Math.PI / 2; // Top

        items.forEach((item: any, i: number) => {
            const angle = startAngle + (i * angleStep);

            const x = centerX + Math.cos(angle) * orbitRadius;
            const y = centerY + Math.sin(angle) * orbitRadius;

            // Draw Connector Arrow from Center to Item (or vice versa)
            // Draw before shape so it's behind
            // From edge of center bubble to item bubble?
            // Simple line from center to x,y is easiest, send to back.
            const conn = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, centerX, centerY, x, y);
            conn.getLineFill().setSolidFill(theme.colors.neutralGray);
            conn.setWeight(2);
            conn.sendToBack();
            // Note: sendToBack acts on the page, might go behind background if not careful. 
            // Usually fine.

            // Satellite Bubble (Number)
            const bubbleR = 50;
            const bubble = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - bubbleR / 2, y - bubbleR / 2, bubbleR, bubbleR);
            bubble.getFill().setSolidFill('#FFFFFF');
            bubble.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            bubble.getBorder().setWeight(3);

            setStyledText(bubble, String(i + 1).padStart(2, '0'), {
                size: 20,
                bold: true,
                color: settings.primaryColor,
                align: SlidesApp.ParagraphAlignment.CENTER
            }, theme);
            try { bubble.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Label (Outside the bubble)
            // Position relative to bubble based on angle?
            // Left items -> Text on Left. Right -> Right.
            const isRight = x > centerX;
            const textW = 120;
            const textH = 60;
            const textX = isRight ? x + bubbleR / 2 + 10 : x - bubbleR / 2 - textW - 10;
            const textY = y - textH / 2;

            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textX, textY, textW, textH);
            const label = item.label || item.title || '';
            setStyledText(textBox, label, {
                size: 14,
                bold: true,
                color: theme.colors.textPrimary,
                align: isRight ? SlidesApp.ParagraphAlignment.START : SlidesApp.ParagraphAlignment.END
            }, theme);
            try { textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        });
    }
}
