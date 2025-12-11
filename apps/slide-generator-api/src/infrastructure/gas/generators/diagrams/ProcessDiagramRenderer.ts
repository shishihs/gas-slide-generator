import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateProcessColors } from '../../../../common/utils/ColorUtils';

export class ProcessDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const steps = data.steps || data.items || [];
        if (!steps.length) return;

        const n = steps.length;
        const startY = area.top + layout.pxToPt(20);
        let currentY = startY;

        // Dynamic Spacing based on count
        const totalH = area.height;
        const itemH = totalH / n;
        // Cap max height per item to keep it tight if few items
        const actualItemH = Math.min(itemH, layout.pxToPt(100));
        const margin = layout.pxToPt(20);

        const numColW = layout.pxToPt(50); // Narrower col for number

        // Horizontal proximity: Bring text closer to number
        const gapNumToText = layout.pxToPt(10);
        const textLeft = area.left + numColW + gapNumToText;
        const textWidth = area.width - (numColW + gapNumToText);

        steps.forEach((step: string, i: number) => {
            const cleanText = String(step || '').replace(/^\s*\d+[\.\s]*/, '');
            const numStr = String(i + 1).padStart(2, '0');

            // 1. Step Number (Large, Styled, Top-aligned with text)
            // Lower Y slightly to align baseline? or alignment TOP is fine if fonts match.
            // Let's keep TOP alignment but ensure sizes balance.

            const numShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, numColW, layout.pxToPt(40));
            // No border/fill
            setStyledText(numShape, numStr, {
                size: 28, // Slightly smaller but Bold
                bold: true,
                color: settings.primaryColor || theme.colors.primary,
                align: SlidesApp.ParagraphAlignment.END // Align right towards the text
            }, theme);
            try { numShape.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

            // 2. Content Text (Closer to Number)
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, currentY + layout.pxToPt(6), textWidth, actualItemH - margin);
            // +6pt offset to align visual top with the large number

            setStyledText(textShape, cleanText, {
                size: 14,
                color: theme.colors.textPrimary,
                align: SlidesApp.ParagraphAlignment.START,
                bold: true // Title-like weight for the step content
            }, theme);
            try { textShape.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

            // 3. Separator (Vertical flow guide)
            // Instead of full width line, how about a small vertical line under the number to lead eye to next?
            // Or keep full width but make it very subtle.
            // Let's stick to full width for "List" feel, but very light.
            if (i < n - 1) {
                const lineY = currentY + actualItemH - margin / 2;
                // Indent line to start at text, keeping numbers in a "gutter"
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, textLeft, lineY, area.left + area.width, lineY);
                line.getLineFill().setSolidFill(theme.colors.faintGray);
                line.setWeight(0.5);
            }

            currentY += actualItemH;
        });
    }
}
