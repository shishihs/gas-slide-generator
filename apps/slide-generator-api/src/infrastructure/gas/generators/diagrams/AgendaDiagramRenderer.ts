import { IDiagramRenderer } from '../IDiagramRenderer';
import { LayoutManager } from '../../../../../common/utils/LayoutManager';
import { SlideUtils } from '../../../../../common/utils/SlideUtils';

export class AgendaDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const items: string[] = data.items || [];
        if (items.length === 0) return;

        // Configuration
        const COLUMN_COUNT = 2;
        const ROW_COUNT = Math.ceil(items.length / COLUMN_COUNT);

        // Use full width of the area, but add padding between items
        const GAP_X = 40;
        const GAP_Y = 25;

        const itemWidth = (area.width - (GAP_X * (COLUMN_COUNT - 1))) / COLUMN_COUNT;
        const itemHeight = Math.min((area.height - (GAP_Y * (ROW_COUNT - 1))) / ROW_COUNT, 80); // Cap height to keep it card-like

        items.forEach((itemText: string, index: number) => {
            const col = index % COLUMN_COUNT;
            const row = Math.floor(index / COLUMN_COUNT);

            const x = area.left + (col * (itemWidth + GAP_X));
            const y = area.top + (row * (itemHeight + GAP_Y));

            // 1. Group container (conceptual)

            // 2. Number Box (Square on the left)
            const numberSize = itemHeight; // Square
            const numberShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, numberSize, numberSize);
            numberShape.getBorder().setTransparent();
            numberShape.getFill().setSolidFill(settings.primaryColor); // Changed from hex to settings.primaryColor

            const numberText = numberShape.getText();
            numberText.setText((index + 1).toString().padStart(2, '0'));
            const numberStyle = numberText.getTextStyle();
            numberStyle.setForegroundColor('#FFFFFF');
            numberStyle.setFontSize(28); // Large number
            numberStyle.setBold(true);
            numberText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

            // 3. Text Box (Right side)
            const textWidth = itemWidth - numberSize;
            const textX = x + numberSize;

            // Background for text (Subtle card)
            const cardShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, textX, y, textWidth, itemHeight);
            cardShape.getFill().setSolidFill(settings.card_bg || '#F5F5F5'); // Very light gray from theme
            cardShape.getBorder().setTransparent();

            // Text content
            const textRange = cardShape.getText();
            textRange.setText(itemText);
            const textStyle = textRange.getTextStyle();
            textStyle.setForegroundColor(settings.text_primary || '#333333');
            textStyle.setFontSize(18); // Readable size

            // Vertical alignment
            textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
            // Center vertically is tricky in GAS, usually done by padding or Shape alignment. 
            // SlidesApp doesn't have 'setVerticalAlignment' easily on shape text, so we rely on padding.

            // Group the item parts
            slide.group([numberShape, cardShape]);
        });
    }
}
