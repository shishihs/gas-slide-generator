import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { addFooter } from '../../../common/utils/SlideUtils';
import * as ColorUtils from '../../../common/utils/ColorUtils';
import { DiagramRendererFactory } from './diagrams/DiagramRendererFactory';

export class GasDiagramSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {
        Logger.log(`[GasDiagramSlideGenerator] Generating Diagram Slide: ${data.layout || data.type}`);
        Logger.log(`[Debug] ColorUtils loaded: ${!!ColorUtils}`); // Ensure dependency is kept by bundler

        // Set Title

        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
        if (titlePlaceholder) {
            try {
                titlePlaceholder.asShape().getText().setText(data.title || '');
            } catch (e) {
                Logger.log(`Warning: Title placeholder found but text could not be set. ${e}`);
            }
        }

        // Priority: JSON type field > layout field
        const type = (data.type || data.layout || '').toLowerCase();
        Logger.log('Generating Diagram Slide: ' + type);

        // Identify the target placeholder to use as the drawing canvas
        // Priority: BODY -> OBJECT -> PICTURE
        const placeholders = slide.getPlaceholders();
        const getPlaceholderTypeSafe = (p: GoogleAppsScript.Slides.PageElement): GoogleAppsScript.Slides.PlaceholderType | null => {
            try {
                const shape = p.asShape();
                if (shape) {
                    return shape.getPlaceholderType();
                }
            } catch (e) {
                // Not a shape or error getting type
            }
            return null;
        };

        const targetPlaceholder = placeholders.find(p => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.BODY)
            || placeholders.find(p => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.OBJECT)
            || placeholders.find(p => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.PICTURE);

        // Get work area with fallback to layout positions
        let workArea: { left: number; top: number; width: number; height: number };
        if (targetPlaceholder) {
            workArea = {
                left: targetPlaceholder.getLeft(),
                top: targetPlaceholder.getTop(),
                width: targetPlaceholder.getWidth(),
                height: targetPlaceholder.getHeight()
            };
        } else {
            // Use type-specific layout area or fallback to contentSlide.body
            const typeKey = `${type}Slide`;
            const areaRect = layout.getRect(`${typeKey}.area`) || layout.getRect('contentSlide.body');
            workArea = areaRect;
        }

        Logger.log(`WorkArea for ${type}: left=${workArea.left}, top=${workArea.top}, width=${workArea.width}, height=${workArea.height}`);

        // Remove the target placeholder to clear the stage for the diagram
        if (targetPlaceholder) {
            try {
                targetPlaceholder.remove();
            } catch (e) {
                Logger.log('Warning: Could not remove target placeholder: ' + e);
            }
        }

        // Get elements before drawing (to identify new ones later)
        const elementsBefore = slide.getPageElements().map(e => e.getObjectId());

        // Delegate to Renderer
        try {
            const renderer = DiagramRendererFactory.getRenderer(type);
            if (renderer) {
                renderer.render(slide, data, workArea, settings, layout);
            } else {
                Logger.log('Diagram logic not implemented for type: ' + type);
            }
        } catch (e) {
            Logger.log(`ERROR in drawing ${type}: ${e}`);
        }

        // Get new elements created during drawing and group them
        // Exclude placeholders (title, subtitle) - only group content/diagram elements
        const newElements = slide.getPageElements().filter(e => {
            // Must be a new element (not existing before drawing)
            if (elementsBefore.includes(e.getObjectId())) return false;

            // Exclude placeholder shapes (title, subtitle)
            try {
                const shape = e.asShape();
                if (shape) {
                    const placeholderType = shape.getPlaceholderType();
                    if (placeholderType === SlidesApp.PlaceholderType.TITLE ||
                        placeholderType === SlidesApp.PlaceholderType.SUBTITLE ||
                        placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
                        return false;
                    }
                }
            } catch (e) {
                // Not a shape or can't determine placeholder type - include it
            }
            return true;
        });

        let generatedGroup: GoogleAppsScript.Slides.PageElement | null = null;

        if (newElements.length > 1) {
            try {
                generatedGroup = slide.group(newElements) as any;
                Logger.log(`Grouped ${newElements.length} content elements for ${type}`);
            } catch (e) {
                Logger.log(`Warning: Could not group elements: ${e}`);
            }
        } else if (newElements.length === 1) {
            generatedGroup = newElements[0];
        }

        // Center the generated content within the work area
        if (generatedGroup) {
            const currentWidth = generatedGroup.getWidth();
            const currentHeight = generatedGroup.getHeight();

            const centerX = workArea.left + (workArea.width - currentWidth) / 2;
            const centerY = workArea.top + (workArea.height - currentHeight) / 2;

            generatedGroup.setLeft(centerX);
            generatedGroup.setTop(centerY);
        }

        addFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }
}
