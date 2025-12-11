
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { SlideTheme } from '../../../common/config/SlideTheme';
import {
    insertImageFromUrlOrFileId,
    setStyledText,
    adjustShapeText_External,
    addFooter
} from '../../../common/utils/SlideUtils';

export class GasSectionSlideGenerator implements ISlideGenerator {
    private sectionCounter = 0;

    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {
        const theme = layout.getTheme();

        // Handle Section Numbering (simplified for state)
        this.sectionCounter++;
        const parsedNum = (() => {
            if (Number.isFinite(data.sectionNo)) {
                return Number(data.sectionNo);
            }
            const m = String(data.title || '').match(/^\s*(\d+)[\.．]/);
            return m ? Number(m[1]) : this.sectionCounter;
        })();
        const num = String(parsedNum).padStart(2, '0');

        const ghostRect = layout.getRect('sectionSlide.ghostNum');
        let ghostImageInserted = false;

        if (imageUpdateOption === 'update') {
            if (data.ghostImageBase64 && ghostRect) {
                try {
                    const imageData = insertImageFromUrlOrFileId(data.ghostImageBase64);
                    // Check if imageData is correct type
                    if (imageData && typeof imageData !== 'string') {
                        const ghostImage = slide.insertImage(imageData as GoogleAppsScript.Base.BlobSource);
                        const imgWidth = ghostImage.getWidth();
                        const imgHeight = ghostImage.getHeight();
                        const scale = Math.min(ghostRect.width / imgWidth, ghostRect.height / imgHeight);
                        const w = imgWidth * scale;
                        const h = imgHeight * scale;
                        ghostImage.setWidth(w).setHeight(h)
                            .setLeft(ghostRect.left + (ghostRect.width - w) / 2)
                            .setTop(ghostRect.top + (ghostRect.height - h) / 2);
                        ghostImageInserted = true;
                    }
                } catch (e) {
                }
            }
        }

        if (!ghostImageInserted && ghostRect && imageUpdateOption === 'update') {
            const ghost = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, ghostRect.left, ghostRect.top, ghostRect.width, ghostRect.height);
            ghost.getText().setText(num);
            const ghostTextStyle = ghost.getText().getTextStyle();
            ghostTextStyle.setFontFamily(theme.fonts.family)
                .setFontSize(theme.fonts.sizes.ghostNum)
                .setBold(true);
            try {
                ghostTextStyle.setForegroundColor(theme.colors.ghostGray);
            } catch (e) {
                ghostTextStyle.setForegroundColor(theme.colors.ghostGray);
            }
            try {
                ghost.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) { }
            ghost.sendToBack();
        }

        // Title Placeholder
        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
        if (titlePlaceholder) {
            try {
                const textRange = titlePlaceholder.asShape().getText();
                textRange.setText(data.title || '');
                textRange.getTextStyle().setBold(true);
            } catch (e) {
                Logger.log(`Warning: Section Title placeholder found but text could not be set. ${e}`);
            }
            // Use manual alignment if needed, but template usually handles it.
            // sectionTitle size is also handled by template.
        } else {
            // Fallback
            const titleRect = layout.getRect('sectionSlide.title');
            const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
            titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            setStyledText(titleShape, data.title, {
                size: theme.fonts.sizes.sectionTitle,
                bold: true,
                align: SlidesApp.ParagraphAlignment.CENTER,
                fontType: 'large'
            }, theme);
        }

        addFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }
}
