
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { RequestFactory } from '../RequestFactory';

export class GasContentSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slideId: string, data: any, layout: LayoutManager, pageNum: number, settings: any): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();

        // 1. Title
        if (data.title) {
            const titleId = slideId + '_TITLE';
            const titleRect = layout.getRect('contentSlide.title') || { left: 30, top: 20, width: 900, height: 60 };

            requests.push(RequestFactory.createShape(slideId, titleId, 'TEXT_BOX', titleRect.left, titleRect.top, titleRect.width, titleRect.height));
            requests.push({ insertText: { objectId: titleId, text: data.title } });
            requests.push(RequestFactory.updateTextStyle(titleId, {
                fontSize: 32,
                bold: true,
                color: settings.primaryColor,
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: 'START' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: 'BOTTOM' as any }, fields: 'contentAlignment' } });
        }

        // 2. Subhead
        let subheadHeight = 0;
        if (data.subhead) {
            const subId = slideId + '_SUBHEAD';
            const subRect = layout.getRect('contentSlide.subhead') || { left: 30, top: 90, width: 900, height: 30 };

            requests.push(RequestFactory.createShape(slideId, subId, 'TEXT_BOX', subRect.left, subRect.top, subRect.width, subRect.height));
            requests.push({ insertText: { objectId: subId, text: data.subhead } });
            requests.push(RequestFactory.updateTextStyle(subId, {
                fontSize: 18,
                bold: true,
                color: theme.colors.neutralGray,
                fontFamily: theme.fonts.family
            }));
            subheadHeight = subRect.height;
        }

        // 3. Body Content (Bullets)
        const points = Array.isArray(data.points) ? data.points : [];
        const columns = data.columns || [];
        const isTwoColumn = (columns.length === 2) || (data.twoColumn);

        const bodyRect = layout.getRect('contentSlide.body') || { left: 30, top: 130, width: 900, height: 380 };
        const topY = bodyRect.top + (data.subhead ? 10 : 0); // adjust for subhead if needed, or assume fixed layout
        // Simplified: use layout manager's body top

        if (isTwoColumn) {
            const gap = 30;
            const colW = (bodyRect.width - gap) / 2;

            const leftPoints = columns[0] || (points.slice(0, Math.ceil(points.length / 2)));
            const rightPoints = columns[1] || (points.slice(Math.ceil(points.length / 2)));

            requests.push(...this.createBulletRequests(slideId, slideId + '_BODY_0', leftPoints, bodyRect.left, topY, colW, bodyRect.height, theme));
            requests.push(...this.createBulletRequests(slideId, slideId + '_BODY_1', rightPoints, bodyRect.left + colW + gap, topY, colW, bodyRect.height, theme));
        } else {
            requests.push(...this.createBulletRequests(slideId, slideId + '_BODY', points, bodyRect.left, topY, bodyRect.width, bodyRect.height, theme));
        }

        return requests;
    }

    private createBulletRequests(slideId: string, objectId: string, points: string[], x: number, y: number, w: number, h: number, theme: any): GoogleAppsScript.Slides.Schema.Request[] {
        const reqs: GoogleAppsScript.Slides.Schema.Request[] = [];
        if (!points || points.length === 0) return reqs;

        // Create Shape
        reqs.push(RequestFactory.createShape(slideId, objectId, 'TEXT_BOX', x, y, w, h));

        // Insert Text
        const fullText = points.join('\n');
        reqs.push({ insertText: { objectId: objectId, text: fullText } });

        // Style Text (Body)
        reqs.push(RequestFactory.updateTextStyle(objectId, {
            fontSize: 18,
            color: theme.colors.textPrimary,
            fontFamily: theme.fonts.family
        }));

        // Apply Bullets
        reqs.push({
            createParagraphBullets: {
                objectId: objectId,
                textRange: { type: 'ALL' },
                bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE'
            }
        });

        // Align Top-Left
        reqs.push({ updateParagraphStyle: { objectId: objectId, style: { alignment: 'START' }, fields: 'alignment' } });
        reqs.push({ updateShapeProperties: { objectId: objectId, shapeProperties: { contentAlignment: 'TOP' as any }, fields: 'contentAlignment' } });

        return reqs;
    }
}
