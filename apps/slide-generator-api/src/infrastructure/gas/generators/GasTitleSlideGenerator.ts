
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { RequestFactory } from '../RequestFactory';

export class GasTitleSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    // API Version
    generate(slideId: string, data: any, layout: LayoutManager, pageNum: number, settings: any): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();

        // 1. Title
        if (data.title) {
            const titleId = slideId + '_TITLE';
            // Use layout manager or hardcoded "Title Slide" layout concept
            // Title typically centered or large
            const titleRect = layout.getRect('titleSlide.title') || { left: 50, top: 200, width: 860, height: 100 };

            requests.push(RequestFactory.createShape(
                slideId, titleId, 'TEXT_BOX',
                titleRect.left, titleRect.top, titleRect.width, titleRect.height
            ));
            requests.push({ insertText: { objectId: titleId, text: data.title } });
            requests.push(RequestFactory.updateTextStyle(titleId, {
                fontSize: 48,
                bold: true,
                color: settings.primaryColor || '#000000',
                fontFamily: theme.fonts.family // Use theme font if available
            }));
            requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: 'BOTTOM' as any }, fields: 'contentAlignment' } });
        }

        // 2. Subtitle / Date
        const text = data.subtitle || data.date;
        if (text) {
            const subId = slideId + '_SUBTITLE';
            const subRect = layout.getRect('titleSlide.subtitle') || { left: 50, top: 320, width: 860, height: 50 };

            requests.push(RequestFactory.createShape(
                slideId, subId, 'TEXT_BOX',
                subRect.left, subRect.top, subRect.width, subRect.height
            ));
            requests.push({ insertText: { objectId: subId, text: text } });
            requests.push(RequestFactory.updateTextStyle(subId, {
                fontSize: 24,
                color: theme.colors.neutralGray || '#666666',
                fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: subId, style: { alignment: 'CENTER' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: subId, shapeProperties: { contentAlignment: 'TOP' as any }, fields: 'contentAlignment' } });
        }

        // Add credit image if needed (omitted for brevity unless requested)

        return requests;
    }
}
