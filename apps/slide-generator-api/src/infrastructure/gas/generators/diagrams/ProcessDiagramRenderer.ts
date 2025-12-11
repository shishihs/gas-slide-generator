import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateProcessColors } from '../../../../common/utils/ColorUtils';

export class ProcessDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const steps = data.steps || data.items || [];
        if (!steps.length) return;

        const n = steps.length;
        let boxHPx, arrowHPx, fontSize;
        if (n <= 2) {
            boxHPx = 100; arrowHPx = 25; fontSize = 16;
        } else if (n === 3) {
            boxHPx = 80; arrowHPx = 20; fontSize = 16;
        } else {
            boxHPx = 65; arrowHPx = 15; fontSize = 14;
        }

        const processColors = generateProcessColors(settings.primaryColor, n);
        const startY = area.top + layout.pxToPt(10);
        let currentY = startY;
        const boxHPt = layout.pxToPt(boxHPx),
            arrowHPt = layout.pxToPt(arrowHPx);
        const headerWPt = layout.pxToPt(120);
        const bodyLeft = area.left + headerWPt;
        const bodyWPt = area.width - headerWPt;

        for (let i = 0; i < n; i++) {
            const cleanText = String(steps[i] || '').replace(/^\s*\d+[\.\s]*/, '');

            // Header
            const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, headerWPt, boxHPt);
            header.getFill().setSolidFill(processColors[i]);
            header.getBorder().setTransparent();
            setStyledText(header, `STEP ${i + 1}`, {
                size: fontSize,
                bold: true,
                color: DEFAULT_THEME.colors.backgroundGray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try { header.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Body
            const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bodyLeft, currentY, bodyWPt, boxHPt);
            body.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
            body.getBorder().setTransparent();

            // Body Text
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bodyLeft + layout.pxToPt(20), currentY, bodyWPt - layout.pxToPt(40), boxHPt);
            setStyledText(textShape, cleanText, { size: fontSize });
            try { textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            currentY += boxHPt;
            if (i < n - 1) {
                const arrowLeft = area.left + headerWPt / 2 - layout.pxToPt(8);
                const arrow = slide.insertShape(SlidesApp.ShapeType.DOWN_ARROW, arrowLeft, currentY, layout.pxToPt(16), arrowHPt);
                arrow.getFill().setSolidFill(DEFAULT_THEME.colors.processArrow || DEFAULT_THEME.colors.ghostGray);
                arrow.getBorder().setTransparent();
                currentY += arrowHPt;
            }
        }
    }
}
