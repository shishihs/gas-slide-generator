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
        const radius = Math.min(area.width, area.height) * 0.35;

        // 1. Central Thin Circle (The Path)
        // Instead of bent arrows, we use a clean circle representing the cycle.
        const circleParams = {
            left: centerX - radius,
            top: centerY - radius,
            width: radius * 2,
            height: radius * 2
        };
        const mainCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, circleParams.left, circleParams.top, circleParams.width, circleParams.height);
        mainCircle.getFill().setTransparent();
        mainCircle.getBorder().getLineFill().setSolidFill(theme.colors.ghostGray); // Very subtle
        mainCircle.getBorder().setWeight(1);
        mainCircle.getBorder().setDashStyle(SlidesApp.DashStyle.DOT); // Dotted for "flow"

        // Center Text
        if (data.centerText) {
            const centerW = radius * 1.2;
            const centerTextBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerX - centerW / 2, centerY - layout.pxToPt(20), centerW, layout.pxToPt(40));
            setStyledText(centerTextBox, data.centerText, {
                size: 18,
                bold: true,
                align: SlidesApp.ParagraphAlignment.CENTER,
                color: theme.colors.primary
            }, theme);
            try { centerTextBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        }

        const count = items.length;
        const angleStep = (2 * Math.PI) / count;
        // Start from top (-PI/2)
        const startAngle = -Math.PI / 2;

        items.forEach((item: any, i: number) => {
            const angle = startAngle + (i * angleStep);

            // Item Position (on the circle)
            const itemX = centerX + Math.cos(angle) * radius;
            const itemY = centerY + Math.sin(angle) * radius;

            // Dot at anchor point (Slightly larger for impact)
            const dotR = layout.pxToPt(10);
            const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, itemX - dotR / 2, itemY - dotR / 2, dotR, dotR);
            dot.getFill().setSolidFill(theme.colors.backgroundWhite);
            dot.getBorder().getLineFill().setSolidFill(theme.colors.primary);
            dot.getBorder().setWeight(2);

            // Text Positioning
            const isRight = itemX > centerX;

            const textW = layout.pxToPt(140);
            const textH = layout.pxToPt(60);
            // Reduced margin to keep text visually connected to anchor
            const margin = layout.pxToPt(10);

            let textLeft = isRight ? (itemX + margin) : (itemX - textW - margin);
            // Adjust for exact center alignment cases (Top/Bottom items)
            if (Math.abs(itemX - centerX) < 5) { textLeft = itemX - textW / 2; }

            const textTop = itemY - textH / 2;

            const labelStr = item.label || '';
            const subLabelStr = item.subLabel || `${String(i + 1).padStart(2, '0')}`;

            // Create separate box for Number to style it boldly
            // Number (Large, Accent) -> Label (Normal)

            // Align logic
            const align = Math.abs(itemX - centerX) < 5 ? SlidesApp.ParagraphAlignment.CENTER : (isRight ? SlidesApp.ParagraphAlignment.START : SlidesApp.ParagraphAlignment.END);

            // Single box approach for now to keep it simple but utilize line break
            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, textTop, textW, textH);

            // To properly style mixed content (Bold Number + Normal Text) in one box often requires advanced API or multiple calls.
            // For now, let's stick to the "Editorial" look where the Title is Bold and Number is small/accented or vice versa.
            // Let's make the Number part of the label but separate line.

            setStyledText(textBox, `${subLabelStr}\n${labelStr}`, {
                size: 14,
                bold: true, // Title is bold
                color: theme.colors.textPrimary,
                align: align
            }, theme);
            // Ideally we'd color the number differently.
            // Since we can't easily mixed-style in mock/helper, use formatting:
            // "01 | Title" ? No, "01\nTitle" is stacked.

            // Let's add a small connector line from dot to text if "Minimal" allows,
            // to bridge the gap strongly.
            if (Math.abs(itemX - centerX) > 5) {
                const lineStart = isRight ? (itemX + dotR / 2) : (itemX - dotR / 2);
                const lineEnd = isRight ? (textLeft) : (textLeft + textW);
                const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineStart, itemY, lineEnd, itemY);
                connector.getLineFill().setSolidFill(theme.colors.primary);
                connector.setWeight(1);
            }
        });
    }
}
