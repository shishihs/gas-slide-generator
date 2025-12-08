
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
            try {
                Logger.log(`Attempting to access template ID: ${templateId}`);
                const templateFile = driveApp.getFileById(templateId);
                const newFile = templateFile.makeCopy(presentation.title);
                Logger.log(`Template copied. New File ID: ${newFile.getId()}`);
                pres = slidesApp.openById(newFile.getId());
            } catch (e: any) {
                Logger.log(`Error accessing/copying template: ${e.toString()}`);
                throw new Error(`Failed to access or copy template with ID: ${templateId}. Ensure the ID is correct and the script has permission to access it. Details: ${e.message}`);
            }
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

            // Determine Layout
            let slideLayout: GoogleAppsScript.Slides.Layout;
            const layouts = pres.getLayouts();

            // Helper to find layout by name (case-insensitive)
            const findLayout = (name: string) => {
                return layouts.find(l => l.getLayoutName().toUpperCase() === name.toUpperCase());
            };

            if (slideModel.layout === 'TITLE') {
                slideLayout = findLayout('TITLE') || layouts[0]; // Fallback to first if not found (usually Title)
            } else if (slideModel.layout === 'SECTION') {
                slideLayout = findLayout('SECTION_HEADER') || findLayout('SECTION ONLY') || layouts[1];
            } else if (slideModel.layout === 'CONTENT' || slideModel.layout === 'AGENDA') {
                slideLayout = findLayout('TITLE_AND_BODY') || layouts[2]; // Default content
            } else {
                slideLayout = findLayout('TITLE_AND_BODY') || layouts[layouts.length - 1]; // Fallback
            }

            // Fallback for safety if somehow no layout matched (shouldn't happen with default templates)
            if (!slideLayout) {
                slideLayout = layouts[0];
            }

            const slide = pres.appendSlide(slideLayout);

            if (slideModel.layout === 'TITLE') {
                // For Title slide, we want to respect the template placeholders, 
                // so we pass a flag or handle it in the generator. 
                // The current generator 'draws' shapes. We need to update the generator first 
                // or parallel. The user implementation plan says update Generator as well.
                titleGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (slideModel.layout === 'SECTION') {
                sectionGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (slideModel.layout === 'CONTENT' || slideModel.layout === 'AGENDA') {
                // Treat Agenda as Content for now
                contentGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else {
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
