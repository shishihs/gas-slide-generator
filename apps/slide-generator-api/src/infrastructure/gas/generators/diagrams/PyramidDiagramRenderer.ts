import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generatePyramidColors } from '../../../../common/utils/ColorUtils';

export class PyramidDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();

        const levels = data.levels || data.items || [];
        if (!levels.length) return;
        const levelsToDraw = levels.slice(0, 5); // Allow up to 5

        const levelHeight = layout.pxToPt(60);
        // Vertical gap
        const levelGap = layout.pxToPt(20);

        const totalHeight = (levelHeight * levelsToDraw.length) + (levelGap * (levelsToDraw.length - 1));
        const startY = area.top + (area.height - totalHeight) / 2;

        // Use a centered list approach for the "Pyramid" feel, where width decreases or increases
        // But for editorial, a clean stack with numbering is better than a literal triangle.
        // Let's do a "Stepped List" where the number indicates hierarchy.

        const contentW = layout.pxToPt(500);
        const centerX = area.left + area.width / 2;
        const startX = centerX - contentW / 2;

        levelsToDraw.forEach((level: any, index: number) => {
            const y = startY + index * (levelHeight + levelGap);
            const isLast = index === levelsToDraw.length - 1;

            // 1. Number / Hierarchy Indicator (01, 02...)
            const numStr = String(index + 1).padStart(2, '0');
            const numW = layout.pxToPt(50);
            const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, startX, y, numW, levelHeight);
            setStyledText(numBox, numStr, {
                size: 32,
                bold: true,
                color: settings.primaryColor,
                align: SlidesApp.ParagraphAlignment.END // Right align number towards content
            }, theme);
            try { numBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

            // 2. Separator line (vertical accent)
            // Move closer to number
            const lineX = startX + numW + layout.pxToPt(10);
            // Height: connect visually but leave breathing room
            const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineX, y + layout.pxToPt(5), lineX, y + levelHeight - layout.pxToPt(5));
            line.getLineFill().setSolidFill(theme.colors.ghostGray);
            line.setWeight(1);

            // 3. Content
            // Closer to line
            const contentX = lineX + layout.pxToPt(15);
            const textW = contentW - (contentX - startX);
            const levelTitle = level.title || `Level ${index + 1}`;
            const levelDesc = level.description || '';

            // Title
            const titleH = layout.pxToPt(25);
            const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y, textW, titleH);
            setStyledText(titleBox, levelTitle.toUpperCase(), {
                size: 14,
                bold: true,
                color: theme.colors.textPrimary,
                align: SlidesApp.ParagraphAlignment.START
            }, theme);

            // Description - Closer to title
            if (levelDesc) {
                const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y + titleH, textW, levelHeight - titleH);
                setStyledText(descBox, levelDesc, {
                    size: 12,
                    color: theme.colors.neutralGray,
                    align: SlidesApp.ParagraphAlignment.START
                }, theme);
                try { descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }
            }
        });
    }
}
