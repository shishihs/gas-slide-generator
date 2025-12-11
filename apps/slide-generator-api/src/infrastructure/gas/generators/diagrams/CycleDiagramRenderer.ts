import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class CycleDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const items = data.items || [];
        if (!items.length) return;

        const textLengths = items.map((item: any) => {
            const labelLength = (item.label || '').length;
            const subLabelLength = (item.subLabel || '').length;
            return labelLength + subLabelLength;
        });
        const maxLength = Math.max(...textLengths);
        const avgLength = textLengths.reduce((sum: number, len: number) => sum + len, 0) / textLengths.length;

        const centerX = area.left + area.width / 2;
        const centerY = area.top + area.height / 2;
        const radiusX = area.width / 3.2;
        const radiusY = area.height / 2.6;

        const maxCardW = Math.min(layout.pxToPt(220), radiusX * 0.8);
        const maxCardH = Math.min(layout.pxToPt(100), radiusY * 0.6);

        let cardW, cardH, fontSize;
        if (maxLength > 25 || avgLength > 18) {
            cardW = Math.min(layout.pxToPt(230), maxCardW); cardH = Math.min(layout.pxToPt(105), maxCardH); fontSize = 13;
        } else if (maxLength > 15 || avgLength > 10) {
            cardW = Math.min(layout.pxToPt(215), maxCardW); cardH = Math.min(layout.pxToPt(95), maxCardH); fontSize = 14;
        } else {
            cardW = layout.pxToPt(200); cardH = layout.pxToPt(90); fontSize = 16;
        }

        if (data.centerText) {
            const centerTextBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerX - layout.pxToPt(100), centerY - layout.pxToPt(50), layout.pxToPt(200), layout.pxToPt(100));
            setStyledText(centerTextBox, data.centerText, { size: 20, bold: true, align: SlidesApp.ParagraphAlignment.CENTER, color: DEFAULT_THEME.colors.textPrimary });
            try { centerTextBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        }

        const positions = [
            { x: centerX + radiusX, y: centerY },
            { x: centerX, y: centerY + radiusY },
            { x: centerX - radiusX, y: centerY },
            { x: centerX, y: centerY - radiusY }
        ];

        // Ensure we only draw as many items as we have (up to 4)
        const itemsToDraw = items.slice(0, 4);

        itemsToDraw.forEach((item: any, i: number) => {
            const pos = positions[i];
            const cardX = pos.x - cardW / 2;
            const cardY = pos.y - cardH / 2;

            const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
            card.getFill().setSolidFill(settings.primaryColor);
            card.getBorder().setTransparent();

            const subLabelText = item.subLabel || `${i + 1}番目`;
            const labelText = item.label || '';
            setStyledText(card, `${subLabelText}\n${labelText}`, { size: fontSize, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });

            try {
                card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
                const textRange = card.getText();
                const subLabelEnd = subLabelText.length;
                if (textRange.asString().length > subLabelEnd) {
                    textRange.getRange(0, subLabelEnd).getTextStyle().setFontSize(Math.max(10, fontSize - 2));
                }
            } catch (e) { }
        });

        // Bent Arrows
        const arrowRadiusX = radiusX * 0.75;
        const arrowRadiusY = radiusY * 0.80;
        const arrowSize = layout.pxToPt(80);
        const arrowPositions = [
            { left: centerX + arrowRadiusX, top: centerY - arrowRadiusY, rotation: 90 },
            { left: centerX + arrowRadiusX, top: centerY + arrowRadiusY, rotation: 180 },
            { left: centerX - arrowRadiusX, top: centerY + arrowRadiusY, rotation: 270 },
            { left: centerX - arrowRadiusX, top: centerY - arrowRadiusY, rotation: 0 }
        ];

        arrowPositions.slice(0, itemsToDraw.length).forEach(pos => {
            const arrow = slide.insertShape(SlidesApp.ShapeType.BENT_ARROW, pos.left - arrowSize / 2, pos.top - arrowSize / 2, arrowSize, arrowSize);
            arrow.getFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
            arrow.getBorder().setTransparent();
            arrow.setRotation(pos.rotation);
        });
    }
}
