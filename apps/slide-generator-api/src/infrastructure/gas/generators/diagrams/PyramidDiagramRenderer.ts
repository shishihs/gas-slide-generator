import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generatePyramidColors } from '../../../../common/utils/ColorUtils';

export class PyramidDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const levels = data.levels || data.items || [];
        if (!levels.length) return;

        // Limit to 4 for visual consistency as per reference, or relax? Reference says slice(0, 4)
        const levelsToDraw = levels.slice(0, 4);

        const levelHeight = layout.pxToPt(70);
        const levelGap = layout.pxToPt(2);
        const totalHeight = (levelHeight * levelsToDraw.length) + (levelGap * (levelsToDraw.length - 1));
        const startY = area.top + (area.height - totalHeight) / 2;
        const pyramidWidth = layout.pxToPt(480);
        const textColumnWidth = layout.pxToPt(400);
        const gap = layout.pxToPt(30);
        const pyramidLeft = area.left;
        const textColumnLeft = pyramidLeft + pyramidWidth + gap;

        const pyramidColors = generatePyramidColors(settings.primaryColor, levelsToDraw.length);
        const baseWidth = pyramidWidth;
        const widthIncrement = baseWidth / levelsToDraw.length;
        const centerX = pyramidLeft + pyramidWidth / 2;

        levelsToDraw.forEach((level: any, index: number) => {
            const levelWidth = baseWidth - (widthIncrement * (levelsToDraw.length - 1 - index));
            const levelX = centerX - levelWidth / 2;
            const levelY = startY + index * (levelHeight + levelGap);

            // Shape
            const levelBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, levelX, levelY, levelWidth, levelHeight);
            levelBox.getFill().setSolidFill(pyramidColors[index]);
            levelBox.getBorder().setTransparent();

            // Inner Text (Title)
            const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, levelX, levelY, levelWidth, levelHeight);
            titleShape.getFill().setTransparent();
            titleShape.getBorder().setTransparent();
            const levelTitle = level.title || `レベル${index + 1}`;
            setStyledText(titleShape, levelTitle, {
                size: DEFAULT_THEME.fonts.sizes.body,
                bold: true,
                color: DEFAULT_THEME.colors.backgroundGray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try { titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Connecting Line
            const connectionStartX = levelX + levelWidth;
            const connectionEndX = textColumnLeft;
            const connectionY = levelY + levelHeight / 2;
            if (connectionEndX > connectionStartX) {
                const connectionLine = slide.insertLine(
                    SlidesApp.LineCategory.STRAIGHT,
                    connectionStartX, connectionY, connectionEndX, connectionY
                );
                connectionLine.getLineFill().setSolidFill('#D0D7DE');
                connectionLine.setWeight(1.5);
            }

            // Description Text
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textColumnLeft, levelY, textColumnWidth, levelHeight);
            textShape.getFill().setTransparent();
            textShape.getBorder().setTransparent();
            const levelDesc = level.description || '';
            let formattedText;
            if (levelDesc.includes('•') || levelDesc.includes('・')) {
                formattedText = levelDesc;
            } else if (levelDesc.includes('\n')) {
                formattedText = levelDesc.split('\n').filter((l: string) => l.trim()).slice(0, 2).map((l: string) => `• ${l.trim()}`).join('\n');
            } else {
                formattedText = levelDesc;
            }
            setStyledText(textShape, formattedText, {
                size: DEFAULT_THEME.fonts.sizes.body - 1,
                align: SlidesApp.ParagraphAlignment.START, // Fixed from LEFT
                color: DEFAULT_THEME.colors.textPrimary,
                bold: true
            });
            try { textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        });
    }
}
