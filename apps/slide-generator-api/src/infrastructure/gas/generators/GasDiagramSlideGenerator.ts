
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { DiagramRendererFactory } from './diagrams/DiagramRendererFactory';
import { RequestFactory } from '../RequestFactory';

export class GasDiagramSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slideId: string, data: any, layout: LayoutManager, pageNum: number, settings: any): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];

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
                fontFamily: layout.getTheme().fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: 'START' }, fields: 'alignment' } });
            requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: 'BOTTOM' as any }, fields: 'contentAlignment' } });
        }


        // 2. Delegate to Renderer
        const type = (data.type || data.layout || '').toLowerCase();
        const renderer = DiagramRendererFactory.getRenderer(type);

        if (renderer) {
            // Calculate work area
            // In API, we can't easily "find placeholders" without a read.
            // Assumption: Use standard Layout area defined in LayoutManager
            const typeKey = `${type}Slide`;
            const areaRect = layout.getRect(`${typeKey}.area`) || layout.getRect('contentSlide.body');
            // Mock area structure that Renderers expect
            const workArea = {
                left: areaRect.left,
                top: areaRect.top,
                width: areaRect.width,
                height: areaRect.height
            };

            const diagramReqs = renderer.render(slideId, data, workArea, settings, layout);
            requests.push(...diagramReqs);
        } else {
            Logger.log('Diagram logic not implemented for type: ' + type);
        }

        return requests;
    }
}
