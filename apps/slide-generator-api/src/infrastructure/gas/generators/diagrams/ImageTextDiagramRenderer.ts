import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { BatchTextStyleUtils } from '../../BatchTextStyleUtils';
import { RequestFactory } from '../../RequestFactory';

export class ImageTextDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const theme = layout.getTheme();
        const imageUrl = data.image; // Must be a publicly accessible URL for Batch API
        const points = data.points || [];

        const gap = layout.pxToPt(20);
        const halfW = (area.width - gap) / 2;

        const isImageLeft = data.imagePosition !== 'right';
        const imgX = isImageLeft ? area.left : area.left + halfW + gap;
        const txtX = isImageLeft ? area.left + halfW + gap : area.left;

        const imgId = slideId + '_IMG';

        // Draw Image (only if URL)
        if (imageUrl && typeof imageUrl === 'string' && imageUrl.startsWith('http')) {
            requests.push(RequestFactory.createImage(slideId, imgId, imageUrl, imgX, area.top, halfW, area.height));
            // Image sizing mode? Default is stretch/fit. Batch API `createImage` with size specified will fit/stretch to that box?
            // Actually `createImage` takes size. It will distort if aspect ratio differs.
            // Ideally we should know aspect ratio. But we don't.
            // Legacy code calculated scale. Batch API can't read image dimensions before insert.
            // We'll rely on Slides API behavior (usually stretches).
            // This is a trade-off of Batch API.
        } else {
            // Placeholder
            const phId = slideId + '_IMG_PH';
            requests.push(RequestFactory.createShape(slideId, phId, 'RECTANGLE', imgX, area.top, halfW, area.height));
            requests.push(RequestFactory.updateShapeProperties(phId, theme.colors.ghostGray, 'TRANSPARENT'));
            requests.push(...BatchTextStyleUtils.setText(slideId, phId, 'Image Placeholder (URL required)', {
                size: 14, align: 'CENTER', color: theme.colors.neutralGray
            }, theme));
            requests.push(RequestFactory.updateShapeProperties(phId, null, null, null, 'MIDDLE'));
        }

        // Draw Text
        const txtId = slideId + '_TXT';
        const textContent = Array.isArray(points) ? points.join('\n') : String(points);
        requests.push(RequestFactory.createShape(slideId, txtId, 'TEXT_BOX', txtX, area.top, halfW, area.height));
        requests.push(...BatchTextStyleUtils.setText(slideId, txtId, textContent, {
            size: theme.fonts.sizes.body,
            color: theme.colors.textPrimary
        }, theme));

        return requests;
    }
}
