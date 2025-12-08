
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { CONFIG } from '../../../common/config/SlideConfig';
import {
    setBackgroundImageFromUrl,
    insertImageFromUrlOrFileId,
    setStyledText,
    adjustShapeText_External,
    addCucFooter,
    offsetRect,
    drawStandardTitleHeader,
    drawSubheadIfAny,
    createContentCushion,
    setBulletsWithInlineStyles,
    renderImagesInArea,
    drawBottomBar,
    normalizeImages,
    setBoldTextSize,
    adjustAreaForSubhead
} from '../../../common/utils/SlideUtils';

export class GasContentSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {
        // NOTE: setMainSlideBackground calls setBackgroundImageFromUrl with CONFIG.BACKGROUND_IMAGES.main
        // We replicate it here
        setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.main, CONFIG.COLORS.background_white, imageUpdateOption);

        const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
        drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings, titleWidthPt, imageUpdateOption);

        const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
            data._subhead_widthPt : null;
        const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead, subheadWidthPt);

        let points = Array.isArray(data.points) ? data.points.slice(0) : [];
        // Agenda logic is handled here in main.ts, but purely for Content Slide, if it IS an agenda, it might be better handled by AgendaSlideGenerator.
        // However, the prompt request is migration. Let's keep the logic or assume Agenda is separate.
        // In main.ts, createContentSlide HANDLES agenda logic if title matches "Agenda".
        // We should probably allow the specific Agenda Generator to handle "Agenda" layout, but if "Content" layout is used for Agenda, we support it.
        // For now, let's keep the original logic simplified.

        // Logic to populate points if agenda
        const isAgenda = /(agenda|アジェンダ|目次|本日お伝えすること)/i.test(String(data.title || ''));
        if (isAgenda && points.length === 0) {
            // Note: In main.ts, it accessed global __SLIDE_DATA_FOR_AGENDA. We might not have that here easily.
            // We'll skip auto-population for now or handle it via a passed in context if really needed.
            points = ['本日の目的', '進め方', '次のアクション']; // Default fallback from main.ts
        }

        const hasImages = Array.isArray(data.images) && data.images.length > 0;
        const isTwo = !!(data.twoColumn || data.columns);

        if ((isTwo && (data.columns || points)) || (!isTwo && points && points.length > 0)) {
            if (isTwo) {
                let L = [], R = [];
                if (Array.isArray(data.columns) && data.columns.length === 2) {
                    L = data.columns[0] || [];
                    R = data.columns[1] || [];
                } else {
                    const mid = Math.ceil(points.length / 2);
                    L = points.slice(0, mid);
                    R = points.slice(mid);
                }
                const baseLeftRect = layout.getRect('contentSlide.twoColLeft');
                const baseRightRect = layout.getRect('contentSlide.twoColRight');
                const adjustedLeftRect = adjustAreaForSubhead(baseLeftRect, data.subhead, layout);
                const adjustedRightRect = adjustAreaForSubhead(baseRightRect, data.subhead, layout);
                const leftRect = offsetRect(adjustedLeftRect, 0, dy);
                const rightRect = offsetRect(adjustedRightRect, 0, dy);
                createContentCushion(slide, leftRect, settings, layout);
                createContentCushion(slide, rightRect, settings, layout);

                const padding = layout.pxToPt(20);
                const leftTextRect = { left: leftRect.left + padding, top: leftRect.top + padding, width: leftRect.width - (padding * 2), height: leftRect.height - (padding * 2) };
                const rightTextRect = { left: rightRect.left + padding, top: rightRect.top + padding, width: rightRect.width - (padding * 2), height: rightRect.height - (padding * 2) };

                const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftTextRect.left, leftTextRect.top, leftTextRect.width, leftTextRect.height);
                const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightTextRect.left, rightTextRect.top, rightTextRect.width, rightTextRect.height);
                setBulletsWithInlineStyles(leftShape, L);
                setBulletsWithInlineStyles(rightShape, R);
                setBoldTextSize(leftShape, 16);
                setBoldTextSize(rightShape, 16);

                try {
                    adjustShapeText_External(leftShape, null);
                } catch (e) { }
                try {
                    adjustShapeText_External(rightShape, null);
                } catch (e) { }
            } else {
                const baseBodyRect = layout.getRect('contentSlide.body');
                const adjustedBodyRect = adjustAreaForSubhead(baseBodyRect, data.subhead, layout);
                const bodyRect = offsetRect(adjustedBodyRect, 0, dy);
                createContentCushion(slide, bodyRect, settings, layout);

                const padding = layout.pxToPt(20);
                const textRect = { left: bodyRect.left + padding, top: bodyRect.top + padding, width: bodyRect.width - (padding * 2), height: bodyRect.height - (padding * 2) };
                const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
                setBulletsWithInlineStyles(bodyShape, points);
                setBoldTextSize(bodyShape, 16);
                try {
                    bodyShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
                } catch (e) { }
                try {
                    adjustShapeText_External(bodyShape, null);
                } catch (e) { }
            }
        }

        if (hasImages && !points.length && !isTwo) {
            const baseArea = layout.getRect('contentSlide.body');
            const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
            const area = offsetRect(adjustedArea, 0, dy);
            createContentCushion(slide, area, settings, layout);
            renderImagesInArea(slide, layout, area, normalizeImages(data.images), imageUpdateOption);
        }

        if (settings.showBottomBar) {
            drawBottomBar(slide, layout, settings);
        }
        addCucFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }
}
