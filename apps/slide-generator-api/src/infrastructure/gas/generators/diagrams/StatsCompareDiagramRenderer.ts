import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateCompareColors } from '../../../../common/utils/ColorUtils';

export class StatsCompareDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const leftTitle = data.leftTitle || '導入前';
        const rightTitle = data.rightTitle || '導入後';
        const stats = data.stats || [];
        if (!stats.length) return;

        const compareColors = generateCompareColors(settings.primaryColor);

        // Header row
        const headerH = layout.pxToPt(45);
        const labelColW = area.width * 0.35;  // Label column width
        const valueColW = (area.width - labelColW) / 2;  // Each value column width

        // Left Title Header
        const leftHeaderX = area.left + labelColW;
        const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftHeaderX, area.top, valueColW, headerH);
        leftHeader.getFill().setSolidFill(compareColors.left);
        leftHeader.getBorder().setTransparent();
        setStyledText(leftHeader, leftTitle, { size: 14, bold: true, color: theme.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER }, theme);
        try { leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        // Right Title Header
        const rightHeaderX = area.left + labelColW + valueColW;
        const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, area.top, valueColW, headerH);
        rightHeader.getFill().setSolidFill(compareColors.right);
        rightHeader.getBorder().setTransparent();
        setStyledText(rightHeader, rightTitle, { size: 14, bold: true, color: theme.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER }, theme);
        try { rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        // Data rows
        const availableHeight = area.height - headerH;
        const rowHeight = Math.min(layout.pxToPt(60), availableHeight / stats.length);
        let currentY = area.top + headerH;

        stats.forEach((stat: any, index: number) => {
            const label = stat.label || '';
            const leftValue = stat.leftValue || '';
            const rightValue = stat.rightValue || '';
            const trend = stat.trend || null;

            // Alternate row background
            const rowBg = index % 2 === 0 ? theme.colors.backgroundGray : '#FFFFFF';

            // Label cell
            const labelCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, labelColW, rowHeight);
            labelCell.getFill().setSolidFill(rowBg);
            labelCell.getBorder().getLineFill().setSolidFill(theme.colors.faintGray);
            setStyledText(labelCell, label, { size: theme.fonts.sizes.body, bold: true, align: SlidesApp.ParagraphAlignment.START }, theme);
            try { labelCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Left value cell
            const leftCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftHeaderX, currentY, valueColW, rowHeight);
            leftCell.getFill().setSolidFill(rowBg);
            leftCell.getBorder().getLineFill().setSolidFill(theme.colors.faintGray);
            setStyledText(leftCell, leftValue, { size: theme.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.left }, theme);
            try { leftCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Right value cell
            const rightCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, currentY, valueColW - (trend ? layout.pxToPt(40) : 0), rowHeight);
            rightCell.getFill().setSolidFill(rowBg);
            rightCell.getBorder().getLineFill().setSolidFill(theme.colors.faintGray);
            setStyledText(rightCell, rightValue, { size: theme.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.right }, theme);
            try { rightCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Trend indicator (optional)
            if (trend) {
                const trendX = rightHeaderX + valueColW - layout.pxToPt(35);
                const trendShape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, trendX, currentY + rowHeight / 4, layout.pxToPt(25), layout.pxToPt(25));
                const isUp = trend.toLowerCase() === 'up';
                const trendColor = isUp ? '#28a745' : '#dc3545';  // Green for up, Red for down
                trendShape.getFill().setSolidFill(trendColor);
                trendShape.getBorder().setTransparent();
                setStyledText(trendShape, isUp ? '↑' : '↓', { size: 12, color: '#FFFFFF', bold: true, align: SlidesApp.ParagraphAlignment.CENTER }, theme);
                try { trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }

            currentY += rowHeight;
        });
    }
}
