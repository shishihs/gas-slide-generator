import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class ProgressDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return;

        const rowH = layout.pxToPt(50);
        const gap = layout.pxToPt(15);
        const startY = area.top + (area.height - (items.length * (rowH + gap))) / 2;

        items.forEach((item: any, i: number) => {
            const y = startY + i * (rowH + gap);
            const labelW = layout.pxToPt(150);
            const barAreaW = area.width - labelW - layout.pxToPt(60);

            // Label
            const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, y, labelW, rowH);
            setStyledText(labelBox, item.label || '', { size: 14, bold: true, align: SlidesApp.ParagraphAlignment.END }, theme);
            try { labelBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Bar BG
            const barBg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW, rowH / 3);
            barBg.getFill().setSolidFill(theme.colors.ghostGray);
            barBg.getBorder().setTransparent();

            // Bar FG
            const percent = Math.min(100, Math.max(0, parseInt(item.percent || 0)));
            if (percent > 0) {
                const barFg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW * (percent / 100), rowH / 3);
                barFg.getFill().setSolidFill(settings.primaryColor);
                barFg.getBorder().setTransparent();
            }

            // Value
            const valBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + labelW + barAreaW + layout.pxToPt(30), y, layout.pxToPt(50), rowH);
            setStyledText(valBox, `${percent}%`, { size: 14, color: theme.colors.neutralGray }, theme);
            try { valBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        });
    }
}
