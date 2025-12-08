
import { ISlideRepository } from '../../domain/repositories/ISlideRepository';
import { Presentation } from '../../domain/model/Presentation';
import { LayoutManager } from '../../common/utils/LayoutManager';
import { GasTitleSlideGenerator } from './generators/GasTitleSlideGenerator';
import { GasSectionSlideGenerator } from './generators/GasSectionSlideGenerator';
import { GasContentSlideGenerator } from './generators/GasContentSlideGenerator';
import { CONFIG } from '../../common/config/SlideConfig';

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

        const templateSlides = pres.getSlides();

        // Initialize LayoutManager
        const pageWidth = pres.getPageWidth();
        const pageHeight = pres.getPageHeight();
        const layoutManager = new LayoutManager(pageWidth, pageHeight);

        // Initialize Generators
        // TODO: Pass actual credit image if we have it (e.g. from user properties or drive)
        const titleGenerator = new GasTitleSlideGenerator(null);
        const sectionGenerator = new GasSectionSlideGenerator(null);
        const contentGenerator = new GasContentSlideGenerator(null);

        // Settings (Mock/Default for now. Should interact with UserProperties or Request)
        const settings = {
            primaryColor: CONFIG.COLORS.primary_color,
            enableGradient: false,
            showTitleUnderline: true,
            showBottomBar: true,
            showDateColumn: true,
            showPageNumber: true,
            ...CONFIG.COLORS
        };

        presentation.slides.forEach((slideModel, index) => {
            // Strategy: Use Rich Generator if available
            // Note: Domain Model 'LayoutType' matches with generator logic

            // Common data mapping
            const commonData = {
                title: slideModel.title.value,
                subtitle: slideModel.subtitle,
                date: new Date().toLocaleDateString(),
                points: slideModel.content.items, // Map content items to points for Content/Agenda
                content: slideModel.content.items,
                // Add images or other specific props if they exist in domain model (currently minimal)
            };

            const slide = pres.appendSlide(slidesApp.PredefinedLayout.BLANK);

            if (slideModel.layout === 'TITLE') {
                titleGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (slideModel.layout === 'SECTION') {
                sectionGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (slideModel.layout === 'CONTENT' || slideModel.layout === 'AGENDA') {
                // Treat Agenda as Content for now, or create separate if logic diverges
                contentGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else {
                // Fallback for types not yet fully implemented in new Generators
                // Use Content Generator as generic fallback? Or keep simplified placeholder logic?
                // Let's use ContentGenerator as generic fallback for now if it works, or just basic TITLE_AND_BODY
                // But since we appended BLANK, we must draw something.
                // Using ContentGenerator is safest "Rich" fallback.
                contentGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            }

            // Also add speaker notes if present (Global handling)
            if (slideModel.notes) {
                const notesPage = slide.getNotesPage();
                notesPage.getSpeakerNotesShape().getText().setText(slideModel.notes);
            }
        });

        // Remove the slides that existed from the original template
        const initialCount = templateSlides.length;
        if (templateId) {
            for (let i = 0; i < initialCount; i++) {
                if (pres.getSlides().length > presentation.slides.length) {
                    pres.getSlides()[0].remove();
                }
            }
        } else {
            // Remove the first blank slide created by .create()
            // Main strategy created new slides, so we remove the original one.
            if (pres.getSlides().length > presentation.slides.length) {
                pres.getSlides()[0].remove();
            }
        }

        return pres.getUrl();
    }
}
