import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText, insertImageFromUrlOrFileId } from '../../../../common/utils/SlideUtils';

export class ImageTextDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const imageUrl = data.image;
        const points = data.points || []; // Text content

        // Define areas
        const gap = layout.pxToPt(20);
        const halfW = (area.width - gap) / 2;

        const isImageLeft = data.imagePosition !== 'right';
        const imgX = isImageLeft ? area.left : area.left + halfW + gap;
        const txtX = isImageLeft ? area.left + halfW + gap : area.left;

        // Draw Image
        if (imageUrl) {
            try {
                // using helper to get blob or null
                const blob = insertImageFromUrlOrFileId(imageUrl);
                let img;
                if (blob) {
                    img = slide.insertImage(blob);
                } else if (typeof imageUrl === 'string' && imageUrl.startsWith('http')) {
                    // Fallback to URL direct insert
                    img = slide.insertImage(imageUrl);
                }

                if (img) {
                    // Simple fit logic:
                    const scale = Math.min(halfW / img.getWidth(), area.height / img.getHeight());
                    const w = img.getWidth() * scale;
                    const h = img.getHeight() * scale;
                    img.setWidth(w).setHeight(h).setLeft(imgX + (halfW - w) / 2).setTop(area.top + (area.height - h) / 2);
                } else {
                    throw new Error("Image insert failed");
                }
            } catch (e) {
                Logger.log('Image insert failed: ' + e);
                // Placeholder
                const ph = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, imgX, area.top, halfW, area.height);
                ph.getFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
                setStyledText(ph, 'Image Placeholder', { align: SlidesApp.ParagraphAlignment.CENTER });
            }
        }

        // Draw Text
        const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, txtX, area.top, halfW, area.height);
        const textContent = points.join('\n');
        setStyledText(textBox, textContent, { size: DEFAULT_THEME.fonts.sizes.body });
    }
}
