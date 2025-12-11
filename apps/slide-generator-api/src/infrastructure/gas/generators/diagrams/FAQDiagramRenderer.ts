import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class FAQDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const items = data.items || data.points || [];
        // Support points as strings ("Q:...", "A:...") or objects
        const parsedItems: any[] = [];
        if (items.length && typeof items[0] === 'string') {
            let currentQ = '';
            items.forEach((str: string) => {
                if (str.startsWith('Q:') || str.startsWith('Q.')) currentQ = str;
                else if (str.startsWith('A:') || str.startsWith('A.')) parsedItems.push({ q: currentQ, a: str });
            });
        } else {
            items.forEach((it: any) => parsedItems.push(it));
        }
        if (!parsedItems.length) return;

        const gap = layout.pxToPt(20);
        const itemH = (area.height - (gap * (parsedItems.length - 1))) / parsedItems.length;

        parsedItems.forEach((item, i) => {
            const y = area.top + i * (itemH + gap);
            const iconSize = layout.pxToPt(40);

            // Q Circle
            const qCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, area.left, y + (itemH - iconSize) / 2, iconSize, iconSize);
            qCircle.getFill().setSolidFill(settings.primaryColor);
            qCircle.getBorder().setTransparent();
            setStyledText(qCircle, 'Q', { size: 18, bold: true, color: '#FFFFFF', align: SlidesApp.ParagraphAlignment.CENTER });
            try { qCircle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Content Box
            const boxLeft = area.left + iconSize + layout.pxToPt(15);
            const boxW = area.width - (iconSize + layout.pxToPt(15));
            const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, boxLeft, y, boxW, itemH);
            box.getFill().setSolidFill('#FFFFFF');
            box.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);

            const qText = (item.q || '').replace(/^[QA][:. ]+/, '');
            const aText = (item.a || '').replace(/^[QA][:. ]+/, '');

            setStyledText(box, `Q. ${qText}\n\nA. ${aText}`, { size: 12, color: DEFAULT_THEME.colors.textPrimary });
        });
    }
}
