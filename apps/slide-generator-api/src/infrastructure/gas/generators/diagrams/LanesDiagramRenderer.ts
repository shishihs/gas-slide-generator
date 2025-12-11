import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText, drawArrowBetweenRects } from '../../../../common/utils/SlideUtils';

export class LanesDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const lanes = data.lanes || [];
        const n = Math.max(1, lanes.length);
        const { laneGapPx, lanePadPx, laneTitleHeightPx, cardGapPx, cardMinHeightPx, cardMaxHeightPx, arrowHeightPx, arrowGapPx } = DEFAULT_THEME.diagram;
        const px = (p: number) => layout.pxToPt(p);

        const laneW = (area.width - px(laneGapPx) * (n - 1)) / n;
        const cardBoxes: any[] = [];

        for (let j = 0; j < n; j++) {
            const lane = lanes[j] || { title: '', items: [] };
            const left = area.left + j * (laneW + px(laneGapPx));

            // Lane Header
            const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, area.top, laneW, px(laneTitleHeightPx));
            lt.getFill().setSolidFill(settings.primaryColor);
            lt.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            setStyledText(lt, lane.title || '', { size: DEFAULT_THEME.fonts.sizes.laneTitle, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
            try { lt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Lane Body
            const laneBodyTop = area.top + px(laneTitleHeightPx);
            const laneBodyHeight = area.height - px(laneTitleHeightPx);
            const laneBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, laneBodyTop, laneW, laneBodyHeight);
            laneBg.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
            laneBg.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.laneBorder);

            const items = Array.isArray(lane.items) ? lane.items : [];
            const rows = Math.max(1, items.length);
            const availH = laneBodyHeight - px(lanePadPx) * 2;
            const idealH = (availH - px(cardGapPx) * (rows - 1)) / rows;
            const cardH = Math.max(px(cardMinHeightPx), Math.min(px(cardMaxHeightPx), idealH));
            const firstTop = laneBodyTop + px(lanePadPx) + Math.max(0, (availH - (cardH * rows + px(cardGapPx) * (rows - 1))) / 2);

            cardBoxes[j] = [];
            for (let i = 0; i < rows; i++) {
                const cardTop = firstTop + i * (cardH + px(cardGapPx));
                const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left + px(lanePadPx), cardTop, laneW - px(lanePadPx) * 2, cardH);
                card.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
                card.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);
                setStyledText(card, items[i] || '', { size: DEFAULT_THEME.fonts.sizes.body });
                try { card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

                cardBoxes[j][i] = {
                    left: left + px(lanePadPx),
                    top: cardTop,
                    width: laneW - px(lanePadPx) * 2,
                    height: cardH
                };
            }
        }

        // Draw Arrows
        const maxRows = Math.max(0, ...cardBoxes.map(a => a ? a.length : 0));
        for (let j = 0; j < n - 1; j++) {
            for (let i = 0; i < maxRows; i++) {
                if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
                    drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], px(arrowHeightPx), px(arrowGapPx), settings);
                }
            }
        }
    }
}
