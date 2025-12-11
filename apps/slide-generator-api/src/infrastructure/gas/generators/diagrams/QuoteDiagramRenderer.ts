import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class QuoteDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const text = data.text || (data.points && data.points[0]) || '';
        const author = data.author || (data.points && data.points[1]) || '';


        // Minimal layout: No background box
        // Use a MASSIVE quote mark transparently in the background for impact

        // Big Quote Mark (Ghost element behind)
        const quoteSize = layout.pxToPt(200);
        const quoteMark = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + layout.pxToPt(20), area.top - layout.pxToPt(40), quoteSize, quoteSize);
        setStyledText(quoteMark, '“', {
            size: 200,
            color: '#F0F0F0', // Very faint
            fontType: 'georgia', // Serif
            bold: true
        }, theme);
        quoteMark.sendToBack(); // Ensure it's behind

        const contentW = area.width * 0.9;
        const contentX = area.left + (area.width - contentW) / 2;
        const textTop = area.top + layout.pxToPt(60);

        // Quote Text
        // Use a nice Serif if possible, otherwise clean Sans
        // Quote Text
        // Use a nice Serif if possible, otherwise clean Sans
        const quoteBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, textTop, contentW, layout.pxToPt(160)); // Reduce reserved height
        setStyledText(quoteBox, text, {
            size: 32,
            bold: false,
            color: theme.colors.textPrimary,
            align: SlidesApp.ParagraphAlignment.CENTER,
            fontType: 'georgia'
        }, theme);
        try { quoteBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch (e) { }

        // Author separating line (small, minimal)
        const lineW = layout.pxToPt(40); // Shorter line
        const lineX = area.left + (area.width - lineW) / 2;
        // Position line closer to text bottom. Since we don't know exact text height, we guess or use relative.
        // Let's assume text takes ~120pt max.
        const lineY = textTop + layout.pxToPt(165);

        const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineX, lineY, lineX + lineW, lineY);
        line.getLineFill().setSolidFill(settings.primaryColor);
        line.setWeight(2);

        if (author) {
            // Author closer to line
            const authorBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, lineY + layout.pxToPt(5), contentW, layout.pxToPt(30));
            setStyledText(authorBox, author, {
                size: 14,
                align: SlidesApp.ParagraphAlignment.CENTER,
                color: theme.colors.neutralGray,
                bold: true // Small but bold
            }, theme);
            try { authorBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }
        }
    }
}
