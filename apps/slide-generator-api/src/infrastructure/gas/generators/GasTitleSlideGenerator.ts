
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { CONFIG } from '../../../common/config/SlideConfig';
import {
    insertImageFromUrlOrFileId,
    setStyledText,
    adjustShapeText_External,
    applyTextStyle,
    // drawBottomBar removed
    // drawCreditImage removed
} from '../../../common/utils/SlideUtils';

export class GasTitleSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {
        // NOTE: We do not set background image here if we want to respect the Master/Layout background.
        // If the user wants to OVERRIDE the template background, we can keep it. 
        // For now, let's assume if a template is used, we respect it. 
        // But the previous code always set it. Let's comment it out or make it optional.
        // setBackgroundImageFromUrl removed

        // Populate Title Placeholder
        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);

        if (titlePlaceholder) {
            Logger.log(`Title Slide: Found Title Placeholder`);
            const shape = titlePlaceholder.asShape();
            shape.getText().setText(data.title || '');
        } else {
            Logger.log('Title Slide: WARNING - No Title Placeholder found on this layout.');
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
        // Logo logic removed

        // Footer / Credit
        // drawBottomBar removed
        // drawCreditImage removed

    }
}

