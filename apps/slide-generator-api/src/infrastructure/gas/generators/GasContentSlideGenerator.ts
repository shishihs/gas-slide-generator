
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
        // NOTE: Commented out manual background setting to respect template.
        // setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.main, CONFIG.COLORS.background_white, imageUpdateOption);

        // Title Placeholder
        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
        if (titlePlaceholder) {
            // Using simple setText to respect template style
            titlePlaceholder.asShape().getText().setText(data.title || '');
        } else {
            // Fallback to manual if no placeholder (e.g. wrong layout)
            // But avoiding hardcoded drawStandardTitleHeader if possible.
            // Let's rely on standard fallback being minimal.
        }

        const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
            data._subhead_widthPt : null;
        // Keep manual Subhead for now as it doesn't map cleanly to standard placeholders.
        // But we should consider if Subtitle placeholder is available.
        // Usually Content slides don't have Subtitle placeholder.
        const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead, subheadWidthPt);


        let points = Array.isArray(data.points) ? data.points.slice(0) : [];
        const isAgenda = /(agenda|アジェンダ|目次|本日お伝えすること)/i.test(String(data.title || ''));
        if (isAgenda && points.length === 0) {
            points = ['本日の目的', '進め方', '次のアクション'];
        }

        const hasImages = Array.isArray(data.images) && data.images.length > 0;
        const isTwo = !!(data.twoColumn || data.columns);

        // Body Placeholders
        const bodies = slide.getPlaceholders().filter(p => (p as any).getPlaceholderType() === SlidesApp.PlaceholderType.BODY);

        if (isTwo && bodies.length >= 2) {
            // If we have at least 2 body placeholders, use them!
            let L = [], R = [];
            if (Array.isArray(data.columns) && data.columns.length === 2) {
                L = data.columns[0] || [];
                R = data.columns[1] || [];
            } else {
                const mid = Math.ceil(points.length / 2);
                L = points.slice(0, mid);
                R = points.slice(mid);
            }
            // Assume bodies[0] is Left, bodies[1] is Right (usually order of creation or x-pos)
            // We might want to sort by Left position to be safe.
            const sortedBodies = bodies.map(p => p.asShape()).sort((a, b) => a.getLeft() - b.getLeft());
            setBulletsWithInlineStyles(sortedBodies[0], L); // Use setBulletsWithInlineStyles for content
            setBulletsWithInlineStyles(sortedBodies[1], R);

        } else if (isTwo) {
            // User wants 2 columns but Layout doesn't have 2 body placeholders.
            // Fallback to manual (Canvas Painting) OR Split single body?
            // Since we are moving to "Template Support", manually drawing boxes on top of a template is messy.
            // But we shouldn't fail.
            // Let's try manual fallback for Two Column IF layout doesn't support it.
            // Original logic for two columns:
            if (Array.isArray(data.columns) && data.columns.length === 2) {
                // ... manual logic ...
            }
            // For brevity in this refactor, let's keep the manual 2-column logic from before ONLY AS FALLBACK.
            // Copying original manual logic:
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
            // We use base rects. Note: layout.getRect relies on hardcoded data, which we wanted to avoid.
            // But as a fallback it's fine.
            const adjustedLeftRect = adjustAreaForSubhead(baseLeftRect, data.subhead, layout);
            const adjustedRightRect = adjustAreaForSubhead(baseRightRect, data.subhead, layout);
            const leftRect = offsetRect(adjustedLeftRect, 0, dy);
            const rightRect = offsetRect(adjustedRightRect, 0, dy);
            // createContentCushion(slide, leftRect, settings, layout); // Optional cushion

            const padding = layout.pxToPt(20);
            const leftTextRect = { left: leftRect.left + padding, top: leftRect.top + padding, width: leftRect.width - (padding * 2), height: leftRect.height - (padding * 2) };
            const rightTextRect = { left: rightRect.left + padding, top: rightRect.top + padding, width: rightRect.width - (padding * 2), height: rightRect.height - (padding * 2) };

            const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftTextRect.left, leftTextRect.top, leftTextRect.width, leftTextRect.height);
            const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightTextRect.left, rightTextRect.top, rightTextRect.width, rightTextRect.height);
            setBulletsWithInlineStyles(leftShape, L);
            setBulletsWithInlineStyles(rightShape, R);
        } else {
            // Single Column case
            if (bodies.length > 0) {
                const bodyShape = bodies[0].asShape();
                // We apply content.
                setBulletsWithInlineStyles(bodyShape, points);
                // We do NOT manual bold/size settings to respect template.
            } else {
                // Fallback: No body placeholder found. Draw manually.
                // This effectively handles "Generic Fallback" for any layout issues.
                const baseBodyRect = layout.getRect('contentSlide.body');
                const adjustedBodyRect = adjustAreaForSubhead(baseBodyRect, data.subhead, layout);
                const bodyRect = offsetRect(adjustedBodyRect, 0, dy);
                // createContentCushion(slide, bodyRect, settings, layout);

                const padding = layout.pxToPt(20);
                const textRect = { left: bodyRect.left + padding, top: bodyRect.top + padding, width: bodyRect.width - (padding * 2), height: bodyRect.height - (padding * 2) };
                const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
                setBulletsWithInlineStyles(bodyShape, points);
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
