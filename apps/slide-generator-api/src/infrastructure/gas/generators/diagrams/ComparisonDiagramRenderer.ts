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
            // 1. Accent Line at top
            const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, area.top, colWidth, layout.pxToPt(3));
            line.getFill().setSolidFill(accentColor);
            line.getBorder().setTransparent();

            // 2. Title (Below line, Large)
            const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, area.top + layout.pxToPt(15), colWidth, headerH);
            setStyledText(titleBox, title, {
                size: 24, // Large
                bold: true,
                color: DEFAULT_THEME.colors.textPrimary,
                align: SlidesApp.ParagraphAlignment.START
            });

            // 3. Items list
            let currentY = area.top + headerH + layout.pxToPt(30);

            items.forEach((itemText: string) => {
                // Small subtle bullet line
                const bulletW = layout.pxToPt(12);
                const bulletH = layout.pxToPt(2);
                const bullet = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, currentY + layout.pxToPt(10), bulletW, bulletH);
                bullet.getFill().setSolidFill(accentColor);
                bullet.getBorder().setTransparent();

                // Text
                const textX = x + bulletW + layout.pxToPt(10);
                const textW = colWidth - (bulletW + layout.pxToPt(10));

                const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textX, currentY, textW, layout.pxToPt(30));
                setStyledText(textBox, itemText, {
                    size: 16,
                    color: DEFAULT_THEME.colors.textPrimary,
                    align: SlidesApp.ParagraphAlignment.START
                });

                currentY += layout.pxToPt(40) + itemSpacing;
            });
        };

        drawColumn(area.left, leftTitle, leftItems, compareColors.left);
        drawColumn(area.left + colWidth + gap, rightTitle, rightItems, compareColors.right);
    }
}
