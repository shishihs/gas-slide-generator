import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateCompareColors } from '../../../../common/utils/ColorUtils';

export class BarCompareDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
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

            // Label
            const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, labelColW, rowHeight);
            setStyledText(labelShape, label, { size: DEFAULT_THEME.fonts.sizes.body, bold: true, align: SlidesApp.ParagraphAlignment.START });
            try { labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Bar area
            const barLeft = area.left + labelColW;
            const barTop = currentY + (rowHeight - (barHeight * 2 + barGap)) / 2;

            // Left bar (Before)
            const leftBarWidth = (leftNum / maxValue) * barAreaW;
            if (leftBarWidth > 0) {
                const leftBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barTop, leftBarWidth, barHeight);
                leftBar.getFill().setSolidFill(compareColors.left);
                leftBar.getBorder().setTransparent();
            }
            // Left label
            const leftLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + leftBarWidth + layout.pxToPt(5), barTop, layout.pxToPt(60), barHeight);
            setStyledText(leftLabel, leftValue, { size: 10, color: compareColors.left });

            // Right bar (After)
            const rightBarWidth = (rightNum / maxValue) * barAreaW;
            if (rightBarWidth > 0) {
                const rightBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barTop + barHeight + barGap, rightBarWidth, barHeight);
                rightBar.getFill().setSolidFill(compareColors.right);
                rightBar.getBorder().setTransparent();
            }
            // Right label
            const rightLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + rightBarWidth + layout.pxToPt(5), barTop + barHeight + barGap, layout.pxToPt(60), barHeight);
            setStyledText(rightLabel, rightValue, { size: 10, color: compareColors.right });

            // Trend indicator (optional)
            if (trend) {
                const trendX = area.left + area.width - trendColW;
                const trendShape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, trendX, currentY + rowHeight / 2 - layout.pxToPt(12), layout.pxToPt(24), layout.pxToPt(24));
                const isUp = trend.toLowerCase() === 'up';
                const trendColor = isUp ? '#28a745' : '#dc3545';
                trendShape.getFill().setSolidFill(trendColor);
                trendShape.getBorder().setTransparent();
                setStyledText(trendShape, isUp ? '↑' : '↓', { size: 12, color: '#FFFFFF', bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
                try { trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }

            // Separator line
            if (index < stats.length - 1) {
                const lineY = currentY + rowHeight;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left, lineY, area.left + area.width, lineY);
                line.getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
                line.setWeight(1);
            }

            currentY += rowHeight;
        });
    }
}
