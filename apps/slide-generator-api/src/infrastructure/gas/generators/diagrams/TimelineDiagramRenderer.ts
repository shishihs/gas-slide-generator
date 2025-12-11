import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateTimelineCardColors } from '../../../../common/utils/ColorUtils';

export class TimelineDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const milestones = data.milestones || data.items || [];
        if (!milestones.length) return;

        const inner = layout.pxToPt(80),
            baseY = area.top + area.height * 0.50;
        const leftX = area.left + inner,
            rightX = area.left + area.width - inner;

        // Draw Line (Minimalist thin dark line)
        const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, leftX, baseY, rightX, baseY);
        line.getLineFill().setSolidFill(theme.colors.neutralGray); // Darker than faintGray for contrast
        line.setWeight(1); // Thin line

        const dotR = layout.pxToPt(8); // Slightly smaller precise dot
        const gap = (milestones.length > 1) ? (rightX - leftX) / (milestones.length - 1) : 0;
        const cardW_pt = layout.pxToPt(160); // Narrower text area

        // New vertical spacing parameters
        const connectorH = layout.pxToPt(20); // Height of the vertical connector from axis to text
        const dateHeight = layout.pxToPt(20);
        const labelHeight = layout.pxToPt(50); // Adjusted for label text
        const dateToLabelGap = layout.pxToPt(5); // Gap between date and label

        milestones.forEach((m: any, i: number) => {
            const x = leftX + gap * i;

            // 1. Dot on Axis (Anchor)
            const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
            dot.getFill().setSolidFill('#FFFFFF'); // White center
            dot.getBorder().getLineFill().setSolidFill(settings.primaryColor || theme.colors.primary); // Use settings color if available
            dot.getBorder().setWeight(2);

            // 2. Vertical Connector (Shorter and clearer)
            // Draws from axis up to the bottom of the date box
            const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, baseY - connectorH, x, baseY - dotR / 2);
            connector.getLineFill().setSolidFill(settings.primaryColor || theme.colors.primary);
            connector.setWeight(1.5); // Increased weight for better visibility

            // Calculate positions for Date and Label (both above the axis)
            const dateTop = baseY - connectorH - dateHeight;
            const labelTop = dateTop - labelHeight - dateToLabelGap;

            // 3. Date Text (Closest to axis anchor)
            const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - cardW_pt / 2, dateTop, cardW_pt, dateHeight);
            dateShape.getFill().setTransparent();
            dateShape.getBorder().setTransparent();
            setStyledText(dateShape, String(m.date || ''), {
                size: 12,
                bold: true,
                color: settings.primaryColor || theme.colors.primary,
                align: SlidesApp.ParagraphAlignment.CENTER
            }, theme);
            try { dateShape.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch (e) { }

            // 4. Label Text (Immediately above Date)
            const labelText = String(m.label || m.state || '');
            const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - cardW_pt / 2, labelTop, cardW_pt, labelHeight);
            labelShape.getFill().setTransparent();
            labelShape.getBorder().setTransparent();

            // Adjust body font size based on length
            let bodyFontSize = 12;
            if (labelText.length > 50) bodyFontSize = 10;

            setStyledText(labelShape, labelText, {
                size: bodyFontSize,
                color: theme.colors.textPrimary, // Stark Black/Gray
                align: SlidesApp.ParagraphAlignment.CENTER
            }, theme);
            try {
                dateShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
                labelShape.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
            } catch (e) { }
        });
    }
}
