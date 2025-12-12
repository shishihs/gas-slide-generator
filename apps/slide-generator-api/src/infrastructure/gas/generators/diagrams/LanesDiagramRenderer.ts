import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class LanesDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const lanes = data.lanes || [];
        const n = Math.max(1, lanes.length);
        const { laneGapPx, lanePadPx, laneTitleHeightPx, cardGapPx, cardMinHeightPx, cardMaxHeightPx, arrowHeightPx, arrowGapPx } = theme.diagram;
        const px = (p: number) => layout.pxToPt(p);

        const laneW = (area.width - px(laneGapPx) * (n - 1)) / n;
        const cardBoxes: any[] = [];



        for (let j = 0; j < n; j++) {
            const lane = lanes[j] || { title: '', items: [] };
            const left = area.left + j * (laneW + px(laneGapPx));
            const laneIdPrefix = slideId + `_LANE_${j}`;

            // Lane Header
            const ltId = laneIdPrefix + '_HEAD';
            requests.push(RequestFactory.createShape(slideId, ltId, 'RECTANGLE', left, area.top, laneW, px(laneTitleHeightPx)));
            requests.push(RequestFactory.updateShapeProperties(ltId, settings.primaryColor, settings.primaryColor, 1));
            requests.push(...BatchTextStyleUtils.setText(slideId, ltId, lane.title || '', {
                size: theme.fonts.sizes.laneTitle, bold: true, color: theme.colors.backgroundGray, align: 'CENTER'
            }, theme));
            requests.push(RequestFactory.updateShapeProperties(ltId, null, null, null, 'MIDDLE'));

            // Lane Body
            const laneBodyTop = area.top + px(laneTitleHeightPx);
            const laneBodyHeight = area.height - px(laneTitleHeightPx);
            const lbId = laneIdPrefix + '_BODY';
            requests.push(RequestFactory.createShape(slideId, lbId, 'RECTANGLE', left, laneBodyTop, laneW, laneBodyHeight));
            requests.push(RequestFactory.updateShapeProperties(lbId, theme.colors.backgroundGray, theme.colors.laneBorder));

            const items = Array.isArray(lane.items) ? lane.items : [];
            const rows = Math.max(1, items.length);
            const availH = laneBodyHeight - px(lanePadPx) * 2;
            const idealH = (availH - px(cardGapPx) * (rows - 1)) / rows;
            const cardH = Math.max(px(cardMinHeightPx), Math.min(px(cardMaxHeightPx), idealH));
            const firstTop = laneBodyTop + px(lanePadPx) + Math.max(0, (availH - (cardH * rows + px(cardGapPx) * (rows - 1))) / 2);

            cardBoxes[j] = [];
            for (let i = 0; i < rows; i++) {
                const cardTop = firstTop + i * (cardH + px(cardGapPx));
                const cardId = laneIdPrefix + `_CARD_${i}`;

                requests.push(RequestFactory.createShape(slideId, cardId, 'ROUND_RECTANGLE', left + px(lanePadPx), cardTop, laneW - px(lanePadPx) * 2, cardH));
                requests.push(RequestFactory.updateShapeProperties(cardId, theme.colors.backgroundGray, theme.colors.cardBorder));
                requests.push(...BatchTextStyleUtils.setText(slideId, cardId, items[i] || '', {
                    size: theme.fonts.sizes.body
                }, theme));
                requests.push(RequestFactory.updateShapeProperties(cardId, null, null, null, 'MIDDLE'));

                cardBoxes[j][i] = {
                    left: left + px(lanePadPx),
                    top: cardTop,
                    width: laneW - px(lanePadPx) * 2,
                    height: cardH
                };
            }
        }

        // Draw Arrows Logic using RequestFactory
        const maxRows = Math.max(0, ...cardBoxes.map(a => a ? a.length : 0));
        for (let j = 0; j < n - 1; j++) {
            for (let i = 0; i < maxRows; i++) {
                if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
                    const boxA = cardBoxes[j][i];
                    const boxB = cardBoxes[j + 1][i];

                    const startX = boxA.left + boxA.width;
                    const startY = boxA.top + boxA.height / 2;
                    const endX = boxB.left;
                    const endY = boxB.top + boxB.height / 2;

                    const arrowId = slideId + `_ARR_${j}_${i}`;
                    requests.push(RequestFactory.createLine(slideId, arrowId, startX, startY, endX, endY));

                    const arrowColor = RequestFactory.toRgbColor(settings.primaryColor) || { red: 0, green: 0, blue: 0 };

                    requests.push({
                        updateLineProperties: {
                            objectId: arrowId,
                            lineProperties: {
                                lineFill: { solidFill: { color: { rgbColor: arrowColor } } },
                                weight: { magnitude: 2, unit: 'PT' },
                                endArrow: 'FILL_ARROW'
                            },
                            fields: 'lineFill,weight,endArrow'
                        }
                    });
                }
            }
        }

        return requests;
    }
}
