import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class KPIDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const items = data.items || [];
        if (!items.length) return;

        const cols = items.length > 4 ? 4 : (items.length || 1);
        // Larger space for editorial feel
        const gap = layout.pxToPt(40);
        const cardW = (area.width - (gap * (cols - 1))) / cols;
        const cardH = layout.pxToPt(180);
        const y = area.top + (area.height - cardH) / 2 + layout.pxToPt(20);

        items.forEach((item: any, i: number) => {
            const x = area.left + i * (cardW + gap);

            // No Background Card - Clean Typography
            // Optional: Left border or separator line
            if (i > 0) {
                // Vertical divider line between KPIs
                const lineH = layout.pxToPt(100);
                const lineY = y + (cardH - lineH) / 2;
                const lineX = x - gap / 2;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineX, lineY, lineX, lineY + lineH);
                line.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
                line.setWeight(1);
            }

            // Label (Top - Small, Uppercase, tracked)
            // Editorial tip: Small but spaced out caps looks professional
            // Label (Top - Small, Uppercase, tracked)
            // Editorial tip: Small but spaced out caps looks professional
            // Bring closer to Value
            const labelH = layout.pxToPt(20);
            const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + layout.pxToPt(5), cardW, labelH);
            setStyledText(labelBox, (item.label || 'METRIC').toUpperCase(), {
                size: 11, // Slightly smaller
                color: DEFAULT_THEME.colors.neutralGray,
                align: SlidesApp.ParagraphAlignment.CENTER,
                bold: true
            });
            try { labelBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch (e) { }

            // Value (Dominant element) - Immediately below label
            const valueH = layout.pxToPt(90);
            const valueBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + labelH, cardW, valueH);
            const valStr = String(item.value || '0');
            // Adaptive font size - Huge for impact
            let fontSize = 72;
            if (valStr.length > 4) fontSize = 60;
            if (valStr.length > 6) fontSize = 48;
            if (valStr.length > 10) fontSize = 36;

            setStyledText(valueBox, valStr, {
                size: fontSize,
                bold: true,
                color: settings.primaryColor,
                align: SlidesApp.ParagraphAlignment.CENTER,
                fontType: 'lato'
            });
            try { valueBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

            // Change/Status (Bottom)
            if (item.change || item.status) {
                const statusH = layout.pxToPt(30);
                const statusBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + labelH + valueH, cardW, statusH);

                let color = DEFAULT_THEME.colors.neutralGray;
                let prefix = '';
                if (item.status === 'good') { color = '#2E7D32'; prefix = '↑ '; } // Darker green for pro look
                if (item.status === 'bad') { color = '#C62828'; prefix = '↓ '; }  // Darker red

                setStyledText(statusBox, prefix + (item.change || ''), {
                    size: 14,
                    bold: true,
                    color: color,
                    align: SlidesApp.ParagraphAlignment.CENTER
                });
            }
        });
    }
}
