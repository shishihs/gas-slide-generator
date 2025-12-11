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

        const gap = layout.pxToPt(30);
        // Estimate height per item based on available space
        const itemH = (area.height - (gap * (parsedItems.length - 1))) / parsedItems.length;

        parsedItems.forEach((item, i) => {
            const y = area.top + i * (itemH + gap);

            // Clean Q text
            const qStr = (item.q || '').replace(/^[QA][:. ]+/, '');
            const aStr = (item.a || '').replace(/^[QA][:. ]+/, '');

            // Q Indicator (Small bold accent)
            const qInd = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, y, layout.pxToPt(30), layout.pxToPt(30));
            setStyledText(qInd, 'Q.', { size: 16, bold: true, color: settings.primaryColor });

            // Q Content - Closer to Q.
            const qBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + layout.pxToPt(30), y, area.width - layout.pxToPt(30), layout.pxToPt(40));
            setStyledText(qBox, qStr, { size: 14, bold: true, color: DEFAULT_THEME.colors.textPrimary });

            // A Indicator (Gray)
            // Positioned below Q - Closer
            const aY = y + layout.pxToPt(30);
            const aInd = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, aY, layout.pxToPt(30), layout.pxToPt(30));
            setStyledText(aInd, 'A.', { size: 16, bold: true, color: DEFAULT_THEME.colors.neutralGray });

            // A Content - Closer to A.
            const aBoxHasHeight = itemH - layout.pxToPt(40);
            if (aBoxHasHeight > 10) {
                const aBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + layout.pxToPt(30), aY, area.width - layout.pxToPt(30), aBoxHasHeight);
                setStyledText(aBox, aStr, { size: 12, color: DEFAULT_THEME.colors.textPrimary });
                try { aBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }
            }

            // Separator Line
            if (i < parsedItems.length - 1) {
                const lineY = y + itemH + gap / 2;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left, lineY, area.left + area.width, lineY);
                line.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
                line.setWeight(0.5);
            }
        });
    }
}
