import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class CardsDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const items = data.items || [];
        if (!items.length) return;
        const type = (data.type || '').toLowerCase();
        const hasHeader = type.includes('headercards');

        const cols = data.columns || Math.min(items.length, 3);
        const rows = Math.ceil(items.length / cols);

        // Editorial spacing: Wider gaps for magazine-like feel
        const gap = layout.pxToPt(30);
        const cardW = (area.width - gap * (cols - 1)) / cols;
        const cardH = (area.height - gap * (rows - 1)) / rows;

        items.forEach((item: any, i: number) => {
            const r = Math.floor(i / cols);
            const c = i % cols;
            const x = area.left + c * (cardW + gap);
            const y = area.top + r * (cardH + gap);

            // Item content parsing
            let title = '';
            let desc = '';
            if (typeof item === 'string') {
                const lines = item.split('\n');
                title = lines[0] || '';
                desc = lines.slice(1).join('\n') || '';
            } else {
                title = item.title || item.label || '';
                desc = item.desc || item.description || item.text || '';
            }

            if (hasHeader) {
                // Editorial Style: Top Bar Header
                // Instead of a full colored box, use a sleek top accent line
                const barH = layout.pxToPt(4);
                const bar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, cardW, barH);
                bar.getFill().setSolidFill(settings.primaryColor);
                bar.getBorder().setTransparent();

                // Numbering (01, 02...) for structure
                // Placed subtly at top right or just above title
                const numStr = String(i + 1).padStart(2, '0');
                const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + layout.pxToPt(10), cardW, layout.pxToPt(20));
                setStyledText(numBox, numStr, {
                    size: 14,
                    bold: true,
                    color: DEFAULT_THEME.colors.neutralGray, // Subtle gray
                    align: SlidesApp.ParagraphAlignment.END
                });

                // Title - Large and bold
                const titleTop = y + layout.pxToPt(10);
                const titleH = layout.pxToPt(40);
                const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleTop, cardW, titleH);
                setStyledText(titleBox, title, {
                    size: 18,
                    bold: true,
                    color: DEFAULT_THEME.colors.textPrimary,
                    align: SlidesApp.ParagraphAlignment.START
                });
                try { titleBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }

                // Body - Use smaller font, lighter color
                const descTop = titleTop + titleH + layout.pxToPt(5);
                const descH = cardH - (descTop - y);
                // Ensure descH is positive
                if (descH > 20) {
                    const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, descTop, cardW, descH);
                    setStyledText(descBox, desc, {
                        size: 13,
                        color: typeof DEFAULT_THEME.colors.textSmallFont === 'string' ? DEFAULT_THEME.colors.textSmallFont : '#424242',
                        align: SlidesApp.ParagraphAlignment.START
                    });
                    try { descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) { }
                }

            } else {
                // Minimal Card (No header bar)
                // Use a small dot accent
                const dotSize = layout.pxToPt(6);
                const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y + layout.pxToPt(9), dotSize, dotSize);
                dot.getFill().setSolidFill(settings.primaryColor);
                dot.getBorder().setTransparent();

                // Title next to dot
                const contentX = x + dotSize + layout.pxToPt(12);
                const contentW = cardW - (dotSize + layout.pxToPt(12));

                const titleH = layout.pxToPt(30);
                const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y, contentW, titleH);
                setStyledText(titleBox, title, {
                    size: 16,
                    bold: true,
                    color: DEFAULT_THEME.colors.textPrimary,
                    align: SlidesApp.ParagraphAlignment.START
                });

                // Body below
                const descTop = y + titleH;
                const descH = cardH - titleH;
                if (descH > 20) {
                    const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, descTop, contentW, descH);
                    setStyledText(descBox, desc, {
                        size: 13,
                        color: typeof DEFAULT_THEME.colors.textSmallFont === 'string' ? DEFAULT_THEME.colors.textSmallFont : '#424242',
                        align: SlidesApp.ParagraphAlignment.START
                    });
                }
            }
        });
    }
}
