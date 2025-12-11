import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateCompareColors } from '../../../../common/utils/ColorUtils';

export class ComparisonDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const leftTitle = data.leftTitle || 'プランA';
        const rightTitle = data.rightTitle || 'プランB';
        const leftItems = data.leftItems || [];
        const rightItems = data.rightItems || [];

        const gap = layout.pxToPt(60);
        const colWidth = (area.width - gap) / 2;

        const compareColors = generateCompareColors(settings.primaryColor);
        const headerH = layout.pxToPt(50);
        const itemSpacing = layout.pxToPt(16);

        // Helper to draw a column (Editorial Style)
        const drawColumn = (
            x: number,
            title: string,
            items: string[],
            accentColor: string
        ) => {
            // 1. Accent Line at top (thicker for impact)
            const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, area.top, layout.pxToPt(60), layout.pxToPt(4));
            line.getFill().setSolidFill(accentColor);
            line.getBorder().setTransparent();

            // 2. Title (Below line, Large)
            const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, area.top + layout.pxToPt(15), colWidth, headerH);
            setStyledText(titleBox, title, {
                size: 28, // Even Larger
                bold: true,
                color: DEFAULT_THEME.colors.textPrimary,
                align: SlidesApp.ParagraphAlignment.START
            });

            // 3. Items list
            // Bring closer to title
            let currentY = area.top + headerH + layout.pxToPt(10);

            items.forEach((itemText: string) => {
                // Minimal Dot - Align with text top approx
                const bulletSize = layout.pxToPt(6);
                const bullet = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, currentY + layout.pxToPt(7), bulletSize, bulletSize);
                bullet.getFill().setSolidFill(DEFAULT_THEME.colors.neutralGray);
                bullet.getBorder().setTransparent();

                // Text - Closer to bullet
                const textX = x + bulletSize + layout.pxToPt(8);
                const textW = colWidth - (bulletSize + layout.pxToPt(8));

                // Estimate height? Just give reasonable buffer.
                const itemH = layout.pxToPt(50);

                const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textX, currentY, textW, itemH);
                setStyledText(textBox, itemText, {
                    size: 16,
                    color: DEFAULT_THEME.colors.textPrimary,
                    align: SlidesApp.ParagraphAlignment.START
                });
                try { textBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

                currentY += layout.pxToPt(25) + itemSpacing;
            });
        };

        drawColumn(area.left, leftTitle, leftItems, compareColors.left);
        drawColumn(area.left + colWidth + gap, rightTitle, rightItems, compareColors.right);

        // Center Separator Line
        const centerX = area.left + colWidth + gap / 2;
        const sepY = area.top + layout.pxToPt(20);
        const sepH = area.height - layout.pxToPt(40);
        const sepLine = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, centerX, sepY, centerX, sepY + sepH);
        sepLine.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
        sepLine.setWeight(1);
    }
}
