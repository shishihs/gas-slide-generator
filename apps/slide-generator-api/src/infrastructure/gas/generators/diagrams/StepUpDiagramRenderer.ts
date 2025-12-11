import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class StepUpDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const items = data.items || [];
        if (!items.length) return;
        const count = items.length;

        const gap = layout.pxToPt(20);
        const stepWidth = (area.width - gap * (count - 1)) / count;

        // Simple stair layout
        const maxStepHeight = area.height;
        const minStepHeight = area.height * 0.4;
        const heightIncrement = (maxStepHeight - minStepHeight) / Math.max(1, count - 1);

        items.forEach((item: any, i: number) => {
            const stepH = minStepHeight + (i * heightIncrement);
            const x = area.left + (i * (stepWidth + gap));
            const y = area.top + (area.height - stepH);

            // Background box (Subtle Gray)
            const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, stepWidth, stepH);
            shape.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
            shape.getBorder().setTransparent();

            // Accent Top Bar
            const topBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, stepWidth, layout.pxToPt(4));
            topBar.getFill().setSolidFill(settings.primaryColor);
            topBar.getBorder().setTransparent();

            // Number (Large watermark-like)
            const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + layout.pxToPt(10), stepWidth, layout.pxToPt(60));
            setStyledText(numBox, String(i + 1).padStart(2, '0'), {
                size: 42,
                bold: true,
                color: DEFAULT_THEME.colors.neutralGray,
                align: SlidesApp.ParagraphAlignment.END
            });

            // Title & Content
            const title = item.title || item.label || '';
            const desc = item.desc || item.description || item.text || '';

            const contentY = y + layout.pxToPt(70);

            const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + layout.pxToPt(10), contentY, stepWidth - layout.pxToPt(20), layout.pxToPt(40));
            setStyledText(titleBox, title, {
                size: 18,
                bold: true,
                color: DEFAULT_THEME.colors.textPrimary,
                align: SlidesApp.ParagraphAlignment.START
            });

            const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + layout.pxToPt(10), contentY + layout.pxToPt(45), stepWidth - layout.pxToPt(20), stepH - layout.pxToPt(120));
            // Ensure height is reasonable
            const descH = stepH - layout.pxToPt(120);
            if (descH > 20) {
                setStyledText(descBox, desc, {
                    size: 14,
                    color: DEFAULT_THEME.colors.textSmallFont,
                    align: SlidesApp.ParagraphAlignment.START
                });
                try { descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }
            }
        });
    }
}
