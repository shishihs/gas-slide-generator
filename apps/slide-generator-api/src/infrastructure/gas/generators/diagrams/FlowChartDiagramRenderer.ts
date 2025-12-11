import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class FlowChartDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const steps = data.steps || data.items || [];
        if (!steps.length) return;

        // Similar to Process but maybe boxes with arrows connected
        // Using a simple left-to-right flow for now
        const count = steps.length;
        const gap = 30;
        const boxWidth = (area.width - (gap * (count - 1))) / count;
        const boxHeight = 80;
        const y = area.top + (area.height - boxHeight) / 2;

        steps.forEach((step: any, i: number) => {
            const x = area.left + i * (boxWidth + gap);
            const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, boxWidth, boxHeight);
            shape.getFill().setSolidFill(theme.colors.backgroundGray);
            shape.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            shape.getBorder().setWeight(2);
            setStyledText(shape, typeof step === 'string' ? step : step.label || '', { size: theme.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER }, theme);
            try { shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            if (i < count - 1) {
                // Arrow
                const ax = x + boxWidth;
                const ay = y + boxHeight / 2;
                const bx = x + boxWidth + gap;
                const by = ay;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, ax, ay, bx, by);
                line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
                line.getLineFill().setSolidFill(theme.colors.neutralGray);
            }
        });
    }
}
