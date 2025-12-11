import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class QuoteDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const text = data.text || (data.points && data.points[0]) || '';
        const author = data.author || (data.points && data.points[1]) || '';

        // Background Cushion
        const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
        bg.getFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
        bg.getBorder().setTransparent();

        // Big Quote Marks
        const quoteMark = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, area.top - layout.pxToPt(20), layout.pxToPt(100), layout.pxToPt(100));
        setStyledText(quoteMark, '“', { size: 120, color: DEFAULT_THEME.colors.ghostGray, font: 'Georgia' });

        const contentW = area.width * 0.8;
        const contentX = area.left + (area.width - contentW) / 2;

        const quoteBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, area.top, contentW, area.height - layout.pxToPt(60));
        setStyledText(quoteBox, text, { size: 28, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.CENTER, font: 'Serif' }); // Usage of Serif if available, else standard
        try { quoteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        if (author) {
            const authorBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, area.top + area.height - layout.pxToPt(60), contentW, layout.pxToPt(40));
            setStyledText(authorBox, `— ${author}`, { size: 16, align: SlidesApp.ParagraphAlignment.END, color: DEFAULT_THEME.colors.neutralGray });
        }
    }
}
