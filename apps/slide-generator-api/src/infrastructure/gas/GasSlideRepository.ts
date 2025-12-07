import { ISlideRepository } from '../../domain/repositories/ISlideRepository';
import { Presentation } from '../../domain/model/Presentation';

// This class acts as an Anti-Corruption Layer (ACL) adaptation
// It translates Domain objects into GAS API calls.
export class GasSlideRepository implements ISlideRepository {

    createPresentation(presentation: Presentation, templateId?: string): string {
        const slidesApp = SlidesApp;
        const driveApp = DriveApp;

        let pres: GoogleAppsScript.Slides.Presentation;

        if (templateId) {
            // 1. Copy Template
            const templateFile = driveApp.getFileById(templateId);
            const newFile = templateFile.makeCopy(presentation.title);
            pres = slidesApp.openById(newFile.getId());
        } else {
            // Fallback: Create Blank
            pres = slidesApp.create(presentation.title);
        }

        const templateSlides = pres.getSlides(); // Pre-existing slides in the template
        const masterSlides = pres.getMasters();  // Master layouts (if needed, but usually we just copy existing slides or use layouts)

        // In this strategy, we assume the template HAS layout slides we can duplicate, 
        // OR we just append new blank slides if no template. 
        // Better strategy for "Template": The template has 'Master' layouts named 'TITLE', 'CONTENT' etc.
        // We will append slides using those layouts.

        // 2. Clear initial "copy" content if we want a fresh start, 
        //    BUT usually a template copy is the base. 
        //    Let's remove all slides from the copy first, preserving masters.
        //    (Actually GAS doesn't easily let you delete all slides if 0 left, so we keep one then delete it later)

        // Strategy B: We append new slides using `pres.appendSlide(layout)`.

        presentation.slides.forEach(slideModel => {
            let layout: GoogleAppsScript.Slides.Layout | undefined;

            // Find layout by name (approximate match)
            const layouts = pres.getLayouts();
            // Try to find a layout that matches the requested type (e.g. "TITLE", "CONTENT")
            // This relies on the Template having named layouts. 
            // If not found, use a default mapping to PredefinedLayouts.

            for (const l of layouts) {
                const name = l.getLayoutName();
                // In GAS, we verify layout names. Often they are random IDs unless renamed.
                // If templateId is NOT provided, we use Predefined.
            }

            // Fallback for custom templates: Use Predefined if templateId is missing
            let safeLayout = slidesApp.PredefinedLayout.TITLE_AND_BODY;

            if (slideModel.layout === 'TITLE') safeLayout = slidesApp.PredefinedLayout.TITLE;
            if (slideModel.layout === 'SECTION') safeLayout = slidesApp.PredefinedLayout.SECTION_HEADER;
            if (slideModel.layout === 'AGENDA') safeLayout = slidesApp.PredefinedLayout.TITLE_AND_BODY; // Fallback

            const slide = pres.appendSlide(safeLayout);

            // 3. Placeholder Replacement Strategy
            // We look for shapes matching placeholders or specific IDs.
            // For standard layouts, we use placeholders.

            // Set Title
            const titlePlaceholder = slide.getPlaceholder(slidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(slidesApp.PlaceholderType.CENTERED_TITLE);
            if (titlePlaceholder) {
                titlePlaceholder.asShape().getText().setText(slideModel.title.value);
            }

            // Set Subtitle
            if (slideModel.subtitle) {
                const subtitlePlaceholder = slide.getPlaceholder(slidesApp.PlaceholderType.SUBTITLE);
                if (subtitlePlaceholder) {
                    subtitlePlaceholder.asShape().getText().setText(slideModel.subtitle);
                }
            }

            // Set Body / Content
            const bodyPlaceholder = slide.getPlaceholder(slidesApp.PlaceholderType.BODY);
            if (bodyPlaceholder) {
                const textRange = bodyPlaceholder.asShape().getText();
                textRange.setText(slideModel.content.items.join('\n'));

                // Apply Bullets
                const paragraphs = textRange.getParagraphs();
                for (let i = 0; i < paragraphs.length; i++) {
                    // Check if setBullet exists (it might not on all types, but Paragraph usually has it)
                    // In GAS types, Paragraph.setBullet() exists.
                    try {
                        (paragraphs[i] as any).setBullet(true);
                    } catch (e) {
                        // Ignore if bullet not supported
                    }
                }
            }

            // Set Notes
            if (slideModel.notes) {
                const notesPage = slide.getNotesPage();
                notesPage.getSpeakerNotesShape().getText().setText(slideModel.notes);
            }
        });

        // Remove the slides that existed from the original template copy (if any)
        // We only want the *newly generated* slides.
        // Be careful not to delete the ones we just added. We added them to the END.
        // So we delete indices 0 to (templateSlides.length - 1).
        if (templateId) {
            const currentSlides = pres.getSlides();
            const initialCount = templateSlides.length;
            // We must remove from valid indices. 
            // If we remove index 0 repeatedly, we remove the first N slides.
            for (let i = 0; i < initialCount; i++) {
                // Check to ensure we don't delete our new slides if something went wrong
                if (pres.getSlides().length > presentation.slides.length) {
                    pres.getSlides()[0].remove();
                }
            }
        } else {
            // If we created from scratch, remove the first blank default slide
            if (pres.getSlides().length > presentation.slides.length) {
                pres.getSlides()[0].remove();
            }
        }

        return pres.getUrl();
    }
}
