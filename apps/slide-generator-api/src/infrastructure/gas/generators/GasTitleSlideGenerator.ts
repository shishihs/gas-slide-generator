
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { CONFIG } from '../../../common/config/SlideConfig';
import {
    setBackgroundImageFromUrl,
    insertImageFromUrlOrFileId,
    setStyledText,
    adjustShapeText_External,
    applyTextStyle,
    drawBottomBar,
    drawCreditImage
} from '../../../common/utils/SlideUtils';

export class GasTitleSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {
        // NOTE: We do not set background image here if we want to respect the Master/Layout background.
        // If the user wants to OVERRIDE the template background, we can keep it. 
        // For now, let's assume if a template is used, we respect it. 
        // But the previous code always set it. Let's comment it out or make it optional.
        // setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.title, CONFIG.COLORS.background_white, imageUpdateOption);

        // Populate Title Placeholder
        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
        if (titlePlaceholder) {
            const shape = titlePlaceholder.asShape();
            // shape.getText().setText(data.title || ''); // Simple setText
            // If we want to support markdown-like styles (bold), we can try setStyledText but without resetting fonts.
            // For now, simple text to guarantee template look.
            shape.getText().setText(data.title || '');
        }

        // Populate Subtitle / Date (Subtitle placeholder often used for subtitle or date)
        const subtitlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
        if (subtitlePlaceholder) {
            if (data.date) {
                subtitlePlaceholder.asShape().getText().setText(data.date);
            } else {
                subtitlePlaceholder.asShape().getText().setText('');
            }
        } else {
            // If template has 'Body' placeholder instead (some title slides do), check that.
            const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
            if (bodyPlaceholder && data.date) {
                bodyPlaceholder.asShape().getText().setText(data.date);
            }
        }

        // Handle Logo if it exists (previous code drew it manually).
        if (imageUpdateOption === 'update' && CONFIG.LOGOS.header) {
            const logoRect = layout.getRect('titleSlide.logo'); // This relies on hardcoded POS_PX.
            try {
                if (CONFIG.LOGOS.header) {
                    const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
                    if (imageData && typeof imageData !== 'string') {
                        const logo = slide.insertImage(imageData as GoogleAppsScript.Base.BlobSource);
                        const aspect = logo.getHeight() / logo.getWidth();
                        logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * aspect);
                    }
                }
            } catch (e) { }
        }

        // Footer / Credit
        if (settings.showBottomBar) {
            drawBottomBar(slide, layout, settings);
        }
        if (this.creditImageBlob) {
            const CREDIT_IMAGE_LINK = 'https://note.com/majin_108';
            drawCreditImage(slide, layout, this.creditImageBlob, CREDIT_IMAGE_LINK);
        }

    }
}

