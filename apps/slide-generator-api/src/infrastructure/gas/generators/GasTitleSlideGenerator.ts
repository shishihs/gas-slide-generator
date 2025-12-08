
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
        setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.title, CONFIG.COLORS.background_white, imageUpdateOption);

        if (imageUpdateOption === 'update') {
            const logoRect = layout.getRect('titleSlide.logo');
            try {
                if (CONFIG.LOGOS.header) {
                    const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
                    if (imageData && typeof imageData !== 'string') {
                        const logo = slide.insertImage(imageData as GoogleAppsScript.Base.BlobSource);
                        const aspect = logo.getHeight() / logo.getWidth();
                        logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * aspect);
                    }
                }
            } catch (e) {
            }
        }

        const titleRect = layout.getRect('titleSlide.title');
        const newTop = (layout.pageH_pt - titleRect.height) / 2;
        const newWidth = titleRect.width + layout.pxToPt(60);
        const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, newTop, newWidth, titleRect.height);
        setStyledText(titleShape, data.title, {
            size: CONFIG.FONTS.sizes.title,
            bold: true,
            fontType: 'large'
        });
        try {
            titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        try {
            const titleText = data.title || '';
            if (titleText.indexOf('\n') === -1) {
                const preCalculatedWidth = (data && typeof data._title_widthPt === 'number') ?
                    data._title_widthPt : null;
                if (preCalculatedWidth !== null && preCalculatedWidth < 900) {
                    adjustShapeText_External(titleShape, preCalculatedWidth);
                } else {
                    adjustShapeText_External(titleShape, null);
                }
            }
        } catch (e) {
        }
        try {
            const titleTextRange = titleShape.getText();
            if (!titleTextRange.isEmpty()) {
                const firstRun = titleTextRange.getRuns()[0];
                if (firstRun) {
                    const currentFontSize = firstRun.getTextStyle().getFontSize();
                    if (currentFontSize === 41) {
                        titleTextRange.getTextStyle().setFontSize(40);
                    }
                }
            }
        } catch (e) {
        }

        if (settings.showDateColumn) {
            const dateRect = layout.getRect('titleSlide.date');
            const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, dateRect.left, dateRect.top, dateRect.width, dateRect.height);
            dateShape.getText().setText(data.date || '');
            applyTextStyle(dateShape.getText(), {
                size: CONFIG.FONTS.sizes.date,
                fontType: 'large'
            });
        }

        if (settings.showBottomBar) {
            drawBottomBar(slide, layout, settings);
        }

        if (this.creditImageBlob) {
            const CREDIT_IMAGE_LINK = 'https://note.com/majin_108';
            drawCreditImage(slide, layout, this.creditImageBlob, CREDIT_IMAGE_LINK);
        }
    }
}
