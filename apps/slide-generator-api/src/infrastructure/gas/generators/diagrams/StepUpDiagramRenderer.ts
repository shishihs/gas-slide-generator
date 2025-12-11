import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class StepUpDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const items = data.items || [];
        if (!items.length) return;
        const count = items.length;

        const gap = layout.pxToPt(20);
        // Reduced max steps to fit clean diagonal
        const validCount = Math.min(count, 5);
        const stepWidth = (area.width - gap * (validCount - 1)) / validCount;
        const stepHeight = area.height / validCount;

        items.slice(0, validCount).forEach((item: any, i: number) => {
            // Diagonal positioning
            // For a 'Step Up', bottom-left to top-right? Or simple ascending bars.
            // Let's do ascending from left to right.

            const x = area.left + (i * (stepWidth + gap));
            // y goes UP as i increases (or down if we want top-left start? No, steps go up)
            // But visuals usually go L->R.
            // We'll stack them such that the last one is highest visually?
            // Actually, "Step Up" usually means progress.
            // Let's keep them distributed L->R, but shift Y up?
            // Or just equal height bars but minimal.

            // Let's do Equal height spacing but Y shifts.
            const totalRise = area.height * 0.6;
            const yStep = totalRise / (validCount - 1 || 1);
            // i=0 is lowest (bottom), i=last is highest (top)
            const y = (area.top + area.height - layout.pxToPt(100)) - (i * yStep);

            // Minimal "Step" Shape: Just a horizontal heavy line and vertical connector?
            // Line (The "Step")
            const lineW = stepWidth;
            const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, lineW, layout.pxToPt(4));
            line.getFill().setSolidFill(settings.primaryColor);
            line.getBorder().setTransparent();

            // Number (Sitting directly on top of the line, bottom-left aligned)
            const numStr = String(i + 1).padEnd(2, '0');
            // Shift Y down so it sits tight on the line
            const numBoxH = layout.pxToPt(40);
            const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y - numBoxH + layout.pxToPt(5), lineW, numBoxH);
            setStyledText(numBox, numStr, {
                size: 32,
                bold: true,
                color: settings.primaryColor,
                align: SlidesApp.ParagraphAlignment.START
            }, theme);
            try { numBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch (e) { }

            // Vertical dotted line to ground (Anchor)
            const dropLine = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, y + layout.pxToPt(2), x, area.top + area.height);
            dropLine.getLineFill().setSolidFill(theme.colors.ghostGray);
            dropLine.setDashStyle(SlidesApp.DashStyle.DOT);
            dropLine.setWeight(1);

            // Title & Desc below line (Hanging)
            // Bring closer to line
            const titleContentY = y + layout.pxToPt(10);

            // Allow text to go down
            const textH = area.top + area.height - titleContentY;

            // Title
            const titleHeight = layout.pxToPt(25);
            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleContentY, stepWidth + gap / 2, titleHeight);

            const title = item.title || item.label || '';
            const desc = item.desc || item.description || item.text || '';

            setStyledText(textBox, `${title.toUpperCase()}`, {
                size: 14,
                bold: true,
                color: theme.colors.textPrimary
            }, theme);
            try { textBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

            // Desc
            if (desc) {
                const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleContentY + titleHeight, stepWidth + gap / 2, textH - titleHeight);
                setStyledText(descBox, desc, {
                    size: 12,
                    color: theme.colors.textSmallFont
                }, theme);
                try { descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }
            }
        });
    }
}
