
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { RequestFactory } from '../RequestFactory';

export class GasSectionSlideGenerator implements ISlideGenerator {
    private sectionCounter = 0;

    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    // API Version
    generate(slideId: string, data: any, layout: LayoutManager, pageNum: number, settings: any): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();

        // 1. Ghost Number (Background) - Create FIRST for Z-order
        this.sectionCounter++;
        const parsedNum = (() => {
            if (Number.isFinite(data.sectionNo)) return Number(data.sectionNo);
            const m = String(data.title || '').match(/^\s*(\d+)[\.．]/);
            return m ? Number(m[1]) : this.sectionCounter;
        })();
        const num = String(parsedNum).padStart(2, '0');

        const ghostRect = layout.getRect('sectionSlide.ghostNum') || { left: 680, top: 320, width: 280, height: 200 };
        const ghostId = slideId + '_GHOST';

        requests.push(RequestFactory.createShape(
            slideId, ghostId, 'TEXT_BOX',
            ghostRect.left, ghostRect.top, ghostRect.width, ghostRect.height
        ));
        requests.push({ insertText: { objectId: ghostId, text: num } });
        requests.push(RequestFactory.updateTextStyle(ghostId, {
            bold: true,
            fontSize: theme.fonts.sizes.ghostNum || 250,
            color: theme.colors.ghostGray,
            fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: ghostId, style: { alignment: 'CENTER' }, fields: 'alignment' } });

        // 2. Title (Foreground)
        const titleId = slideId + '_TITLE';
        if (data.title) {
            const titleRect = layout.getRect('sectionSlide.title') || { left: 40, top: 200, width: 700, height: 100 };

            requests.push(RequestFactory.createShape(
                slideId, titleId, 'TEXT_BOX',
                titleRect.left, titleRect.top, titleRect.width, titleRect.height
            ));
            requests.push({ insertText: { objectId: titleId, text: data.title } });
            requests.push(RequestFactory.updateTextStyle(titleId, {
                bold: true,
                fontSize: theme.fonts.sizes.sectionTitle || 52,
                color: theme.colors.primary,
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: 'START' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: 'BOTTOM' as any }, fields: 'contentAlignment' } });
        }

        return requests;
    }
}
