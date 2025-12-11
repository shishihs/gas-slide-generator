import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class KPIDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const items = data.items || [];
        if (!items.length) return;

        const cols = items.length > 4 ? 4 : (items.length || 1);
        const gap = layout.pxToPt(20);
        const cardW = (area.width - (gap * (cols - 1))) / cols;
        const cardH = layout.pxToPt(160);
        const y = area.top + (area.height - cardH) / 2;

        items.forEach((item: any, i: number) => {
            const x = area.left + i * (cardW + gap);

            // Card BG
            const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, cardW, cardH);
            card.getFill().setSolidFill('#FFFFFF');
            card.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);

            const padding = layout.pxToPt(10);

            // Label (Top)
            const labelH = layout.pxToPt(30);
            const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + padding, cardW - padding * 2, labelH);
            setStyledText(labelBox, item.label || 'Metric', { size: 14, color: DEFAULT_THEME.colors.neutralGray, align: SlidesApp.ParagraphAlignment.CENTER });

            // Value (Middle - Large)
            const valueH = layout.pxToPt(70);
            const valueBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + labelH + padding, cardW - padding * 2, valueH);
            const valStr = String(item.value || '0');
            // Adaptive font size
            let fontSize = 48;
            if (valStr.length > 4) fontSize = 36;
            if (valStr.length > 6) fontSize = 28;
            if (valStr.length > 10) fontSize = 24;

            setStyledText(valueBox, valStr, { size: fontSize, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.CENTER });

            // Change/Status (Bottom)
            if (item.change || item.status) {
                const statusH = layout.pxToPt(30);
                const statusBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + labelH + valueH + padding, cardW - padding * 2, statusH);

                let color = DEFAULT_THEME.colors.neutralGray;
                let prefix = '';
                if (item.status === 'good') { color = '#28a745'; prefix = '▲ '; }
                if (item.status === 'bad') { color = '#dc3545'; prefix = '▼ '; }

                setStyledText(statusBox, prefix + (item.change || ''), { size: 14, bold: true, color: color, align: SlidesApp.ParagraphAlignment.CENTER });
            }
        });
    }
}
