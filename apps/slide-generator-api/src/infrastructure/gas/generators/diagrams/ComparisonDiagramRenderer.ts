import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateCompareColors } from '../../../../common/utils/ColorUtils';

export class ComparisonDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
        const leftTitle = data.leftTitle || 'プランA';
        const rightTitle = data.rightTitle || 'プランB';
        const leftItems = data.leftItems || [];
        const rightItems = data.rightItems || [];

        render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
            const theme = layout.getTheme();
            const leftTitle = data.leftTitle || 'プランA';
            const rightTitle = data.rightTitle || 'プランB';
            const leftItems = data.leftItems || [];
            const rightItems = data.rightItems || [];

            const gap = layout.pxToPt(40);
            // Slightly reduced gap to give more space to cards
            const colWidth = (area.width - gap) / 2;
            const headerH = layout.pxToPt(60);

            // Helper to draw a Card Column
            const drawCardColumn = (
                x: number,
                title: string,
                items: string[],
                isPrimary: boolean // Maybe distinguish left/right styles? Or both equal. Let's make both Primary for now.
            ) => {
                // 1. Header Box
                const headerBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, area.top, colWidth, headerH);
                headerBox.getFill().setSolidFill(settings.primaryColor);
                headerBox.getBorder().setTransparent();

                setStyledText(headerBox, title, {
                    size: 24,
                    bold: true,
                    color: '#FFFFFF',
                    align: SlidesApp.ParagraphAlignment.CENTER
                }, theme);
                try { headerBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

                // 2. Body Box (Background for items)
                // Calculate height based on items? Or extend to bottom?
                // Let's extend to near bottom.
                const bodyH = area.height - headerH;
                const bodyY = area.top + headerH;

                const bodyBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, bodyY, colWidth, bodyH);
                bodyBox.getFill().setSolidFill('#F8F9FA'); // Very light gray from theme
                bodyBox.getBorder().getLineFill().setSolidFill(settings.primaryColor);
                bodyBox.getBorder().setWeight(1);

                // 3. Items list
                let currentY = bodyY + layout.pxToPt(20);
                const itemGap = layout.pxToPt(15);

                items.forEach((itemText: string) => {
                    // Icon (Checkmark or Dot)
                    // Using simple shape or text character "✔"
                    const iconSize = 24;
                    const iconBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + 20, currentY, iconSize, iconSize);
                    setStyledText(iconBox, "✔", {
                        size: 18,
                        color: settings.primaryColor,
                        bold: true,
                        align: SlidesApp.ParagraphAlignment.CENTER
                    }, theme);

                    // Text
                    const textX = x + 20 + iconSize + 10;
                    const textW = colWidth - (20 + iconSize + 10 + 20);
                    const itemH = layout.pxToPt(40); // estimate

                    const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textX, currentY, textW, itemH);
                    setStyledText(textBox, itemText, {
                        size: 16,
                        color: theme.colors.textPrimary,
                        align: SlidesApp.ParagraphAlignment.START
                    }, theme);
                    try { textBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

                    currentY += itemH + itemGap;
                });
            };

            drawCardColumn(area.left, leftTitle, leftItems, true);
            drawCardColumn(area.left + colWidth + gap, rightTitle, rightItems, true);
        }
    }
}
