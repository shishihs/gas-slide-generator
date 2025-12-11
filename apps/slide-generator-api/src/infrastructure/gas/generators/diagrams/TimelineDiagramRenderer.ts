import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateTimelineCardColors } from '../../../../common/utils/ColorUtils';

export class TimelineDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const milestones = data.milestones || data.items || [];
        if (!milestones.length) return;

        const centerY = area.top + area.height / 2;
        const startX = area.left + 50;
        const endX = area.left + area.width - 50;
        const usableWidth = endX - startX;

        // 1. Central Axis
        const axis = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, startX, centerY, endX, centerY);
        axis.getLineFill().setSolidFill(theme.colors.neutralGray);
        axis.setWeight(4); // Thick axis

        // 2. Milestones
        const count = milestones.length;
        const gap = usableWidth / (count + 1);

        milestones.forEach((m: any, i: number) => {
            const x = startX + (gap * (i + 1));

            // Alternating Top/Bottom
            const isTop = i % 2 === 0;
            const cardH = 80;
            const cardW = 140;
            const connLen = 40;

            // Connector Y
            const connY1 = centerY;
            const connY2 = isTop ? centerY - connLen : centerY + connLen;

            // Dot on Axis
            const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - 8, centerY - 8, 16, 16);
            dot.getFill().setSolidFill(settings.primaryColor);
            dot.getBorder().setTransparent(); // Clean look

            // Connector Line
            const conn = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, connY1, x, connY2);
            conn.getLineFill().setSolidFill(theme.colors.neutralGray);
            conn.setWeight(2);

            // Card
            const cardY = isTop ? connY2 - cardH : connY2;
            const cardX = x - cardW / 2;

            // Shadow box (Offset grey)
            const shadow = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX + 3, cardY + 3, cardW, cardH);
            shadow.getFill().setSolidFill('#E0E0E0');
            shadow.getBorder().setTransparent();

            // Main Card
            const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
            card.getFill().setSolidFill('#FFFFFF');
            card.getBorder().getLineFill().setSolidFill(theme.colors.ghostGray);
            card.getBorder().setWeight(1);

            // Content
            const date = m.date || m.year || '';
            const label = m.label || m.title || m.text || '';

            const textRange = card.getText();
            textRange.setText(`${date}\n${label}`);

            // Style Date (Top line)
            if (date) {
                const dateRange = textRange.getRange(0, date.length);
                dateRange.getTextStyle().setBold(true).setForegroundColor(settings.primaryColor).setFontSize(14);
            }

            // Style Label
            const labelStart = date ? date.length + 1 : 0;
            if (label.length > 0) {
                const labelRange = textRange.getRange(labelStart, textRange.getLength() - labelStart);
                labelRange.getTextStyle().setBold(false).setForegroundColor(theme.colors.textPrimary).setFontSize(12);
            }

            // Alignment
            textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
            try { card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Group for safety? No, groups are hard to manage dynamically sometimes.
        });
    }
}
