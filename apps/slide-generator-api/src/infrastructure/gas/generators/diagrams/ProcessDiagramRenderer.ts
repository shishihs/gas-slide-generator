import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateProcessColors } from '../../../../common/utils/ColorUtils';

export class ProcessDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const steps = data.steps || data.items || [];
        if (!steps.length) return;

        const count = steps.length;
        // Layout: Horizontal flow? Or Vertical?
        // Process usually Horizontal L->R. User request "Chevron Layout" implies widely used horizontal bars.
        // Let's do Horizontal flow if it fits, else Vertical list with arrows.
        // Given text length might be long, Vertical List with "Chevron-like" headers might be safer?
        // But "Process" usually implies timeline-ish.
        // Let's stick to Vertical List but make it look like a "Chain".

        // Vertical Chain Layout
        const itemHeight = Math.min(area.height / count, 120);
        const gap = 10;
        const boxWidth = area.width;

        steps.forEach((step: string, i: number) => {
            const y = area.top + (i * itemHeight);
            const h = itemHeight - gap;

            // 1. Pentagon/Chevron Shape for the container? 
            // Let's use a Rounded Rectangle with a heavy left border or number circle.

            // Container Box
            const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left, y, boxWidth, h);
            box.getFill().setSolidFill(settings.card_bg || '#F8F9FA');
            box.getBorder().setTransparent();

            // 2. Number Circle (Left)
            const circleSize = 40;
            const circleX = area.left + 20;
            const circleY = y + (h - circleSize) / 2;

            const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, circleX, circleY, circleSize, circleSize);
            circle.getFill().setSolidFill(settings.primaryColor);
            circle.getBorder().setTransparent();

            // Number
            setStyledText(circle, String(i + 1), {
                size: 18,
                bold: true,
                color: '#FFFFFF',
                align: SlidesApp.ParagraphAlignment.CENTER
            }, theme);
            try { circle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // 3. Arrow Down (between items)? Only if not last.
            if (i < count - 1) {
                // Draw a small arrow pointing down from this box to next?
                // Or just the boxes themselves imply flow.
            }

            // 4. Content
            const textX = circleX + circleSize + 20;
            const textW = boxWidth - (textX - area.left) - 20;

            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textX, y + 5, textW, h - 10);
            const cleanText = String(step || '').replace(/^\s*\d+[\.\s]*/, '');

            setStyledText(textBox, cleanText, {
                size: 14,
                color: theme.colors.textPrimary,
                align: SlidesApp.ParagraphAlignment.START,
                bold: false
            }, theme);
            // Vertically center text
            try { textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        });
    }
}
```
