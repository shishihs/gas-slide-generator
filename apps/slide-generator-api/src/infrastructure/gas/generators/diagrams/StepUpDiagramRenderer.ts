import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class StepUpDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return;
        const count = items.length;
        const validCount = Math.min(count, 5);

        const gap = 20;
        const barWidth = (area.width - (gap * (validCount - 1))) / validCount;

        // Max height for the tallest bar
        const maxBarH = area.height * 0.7;
        const minBarH = area.height * 0.2;

        items.slice(0, validCount).forEach((item: any, i: number) => {
            // Linear growth
            const ratio = i / (validCount - 1 || 1);
            const barH = minBarH + (maxBarH - minBarH) * ratio;

            const x = area.left + (i * (barWidth + gap));
            const y = area.top + area.height - barH;

            // 1. Bar Shape
            const bar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, barWidth, barH);
            bar.getFill().setSolidFill(settings.primaryColor);
            bar.getBorder().setTransparent();

            // 2. Number (Inside bar, bottom)
            const numStr = String(i + 1).padEnd(2, '0');
            const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + barH - 40, barWidth, 30);
            setStyledText(numBox, numStr, {
                size: 24,
                bold: true,
                color: '#FFFFFF',
                align: SlidesApp.ParagraphAlignment.CENTER
            }, theme);

            // 3. Title (Above Bar)
            const titleContent = item.title || item.label || '';
            const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y - 40, barWidth, 35);
            setStyledText(titleBox, titleContent, {
                size: 12, // slightly smaller to fit
                bold: true,
                color: theme.colors.textPrimary,
                align: SlidesApp.ParagraphAlignment.CENTER
            }, theme);
            try { titleBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch (e) { }

            // 4. Description (Inside Bar? or Below? or Hover?)
            // If we have description, maybe put it inside the bar at top (white text) if bar is tall enough?
            // Or put everything below the bar? Use User Request "Information enumeration is weak".
            // Let's put Description inside the bar (top aligned, white) if space permits.
            if (item.desc || item.description) {
                const desc = item.desc || item.description;
                const descBoxH = barH - 50; // space above number
                if (descBoxH > 40) {
                    const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + 5, y + 5, barWidth - 10, descBoxH);
                    setStyledText(descBox, desc, {
                        size: 11,
                        color: '#FFFFFF', // White text on Primary Color
                        align: SlidesApp.ParagraphAlignment.CENTER
                    }, theme);
                    try { descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }
                }
            }
        });
    }
}
