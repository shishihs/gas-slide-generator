import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateCompareColors } from '../../../../common/utils/ColorUtils';

export class BarCompareDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const leftTitle = data.leftTitle || '導入前';
        const rightTitle = data.rightTitle || '導入後';
        const stats = data.stats || [];
        if (!stats.length) return;

        const compareColors = generateCompareColors(settings.primaryColor);

        // Find max value for scaling
        let maxValue = 0;
        stats.forEach((stat: any) => {
            const leftNum = parseFloat(String(stat.leftValue || '0').replace(/[^0-9.]/g, '')) || 0;
            const rightNum = parseFloat(String(stat.rightValue || '0').replace(/[^0-9.]/g, '')) || 0;
            maxValue = Math.max(maxValue, leftNum, rightNum);
        });
        if (maxValue === 0) maxValue = 100;  // Fallback

        // Layout
        const labelColW = area.width * 0.2;
        const barAreaW = area.width * 0.6;
        const valueColW = area.width * 0.1;
        const trendColW = area.width * 0.1;

        const rowHeight = Math.min(layout.pxToPt(80), area.height / stats.length);
        const barHeight = layout.pxToPt(18);
        const barGap = layout.pxToPt(4);
        let currentY = area.top;

        stats.forEach((stat: any, index: number) => {
            const label = stat.label || '';
            const leftValue = stat.leftValue || '';
            const rightValue = stat.rightValue || '';
            const trend = stat.trend || null;

            const leftNum = parseFloat(String(leftValue).replace(/[^0-9.]/g, '')) || 0;
            const rightNum = parseFloat(String(rightValue).replace(/[^0-9.]/g, '')) || 0;

            // Label (Left Column)
            const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, labelColW, rowHeight);
            setStyledText(labelShape, label, {
                size: 14,
                bold: true,
                align: SlidesApp.ParagraphAlignment.START,
                color: theme.colors.textPrimary
            }, theme);
            try { labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Bar Area
            // Bar Area
            const barLeft = area.left + labelColW + layout.pxToPt(20);
            const maxBarWidth = area.width - (labelColW + layout.pxToPt(80)); // Leave room on right for diff

            const barGap = layout.pxToPt(2); // Tighter gap
            const barY = currentY + (rowHeight - (barHeight * 2 + barGap * 2)) / 2;

            // Minimal Bars
            // Left (Before) - Gray, Thin
            const leftBarW = (leftNum / maxValue) * maxBarWidth;
            if (leftBarW > 0) {
                const b = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY, leftBarW, barHeight);
                b.getFill().setSolidFill(theme.colors.neutralGray);
                b.getBorder().setTransparent();
                // Value Text inside or end? End looks cleaner.
                const v = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + leftBarW + layout.pxToPt(5), barY - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5));
                setStyledText(v, leftValue, { size: 10, color: theme.colors.neutralGray }, theme);
            }

            // Right (After) - Color, Slightly Thicker/Same
            const rightBarW = (rightNum / maxValue) * maxBarWidth;
            if (rightBarW > 0) {
                // Closer to left bar
                const b = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY + barHeight + barGap, rightBarW, barHeight);
                b.getFill().setSolidFill(settings.primaryColor);
                b.getBorder().setTransparent();
                const v = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + rightBarW + layout.pxToPt(5), barY + barHeight + barGap - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5));
                setStyledText(v, rightValue, { size: 10, bold: true, color: settings.primaryColor }, theme);
            }

            // Trend Arrow on far right (Text only, no ball)
            if (trend) {
                const trendX = area.left + area.width - layout.pxToPt(40);
                const isUp = trend.toLowerCase() === 'up';
                const trendShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, trendX, currentY, layout.pxToPt(40), rowHeight);
                const color = isUp ? '#2E7D32' : '#C62828';
                const sym = isUp ? '↑' : '↓';
                setStyledText(trendShape, sym, {
                    size: 20,
                    bold: true,
                    color: color,
                    align: SlidesApp.ParagraphAlignment.CENTER
                }, theme);
                try { trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }

            // Separator line (clean divider)
            if (index < stats.length - 1) {
                const lineY = currentY + rowHeight;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left, lineY, area.left + area.width, lineY);
                line.getLineFill().setSolidFill(theme.colors.ghostGray);
                line.setWeight(0.5); // Very thin
            }

            currentY += rowHeight;
        });
    }
}
