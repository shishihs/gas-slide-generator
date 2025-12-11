
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import {
    insertImageFromUrlOrFileId,
    setStyledText,
    adjustShapeText_External,
    applyTextStyle
} from '../../../common/utils/SlideUtils';

export class GasTitleSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {


        // Populate Title Placeholder
        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);

        if (titlePlaceholder) {
            if (data.title) {
                try {
                    titlePlaceholder.asShape().getText().setText(data.title);
                } catch (e) {
                    Logger.log(`Warning: Title placeholder found but text could not be set. ${e}`);
                }
            } else {
                titlePlaceholder.remove();
            }
        } else {
            Logger.log('Title Slide: WARNING - No Title Placeholder found on this layout.');
        }

        // Populate Subtitle / Date (Subtitle placeholder often used for subtitle or date)
        const subtitlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
        if (subtitlePlaceholder) {
            if (data.date) {
                try {
                    subtitlePlaceholder.asShape().getText().setText(data.date);
                } catch (e) { }
            } else {
                try {
                    subtitlePlaceholder.asShape().getText().setText('');
                } catch (e) { }
            }
        } else {
            // If template has 'Body' placeholder instead (some title slides do), check that.
            const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
            if (bodyPlaceholder && data.date) {
                try {
                    bodyPlaceholder.asShape().getText().setText(data.date);
                } catch (e) { }
            }
        }



    }
}

