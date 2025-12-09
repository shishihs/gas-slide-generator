import { ISlideRepository } from '../../domain/repositories/ISlideRepository';
import { Presentation } from '../../domain/model/Presentation';
import { LayoutManager } from '../../common/utils/LayoutManager';
import { GasTitleSlideGenerator } from './generators/GasTitleSlideGenerator';
import { GasSectionSlideGenerator } from './generators/GasSectionSlideGenerator';
import { GasContentSlideGenerator } from './generators/GasContentSlideGenerator';
import { GasDiagramSlideGenerator } from './generators/GasDiagramSlideGenerator';
import { CONFIG } from '../../common/config/SlideConfig';

// This class acts as an Anti-Corruption Layer (ACL) adaptation
// It translates Domain objects into GAS API calls.
export class GasSlideRepository implements ISlideRepository {

    createPresentation(presentation: Presentation, templateId?: string, destinationId?: string): string {
        const slidesApp = SlidesApp;
        const driveApp = DriveApp;

        let pres: GoogleAppsScript.Slides.Presentation;

        if (destinationId) {
            // 1. Use Existing Destination (Pre-copied by Caller)
            try {
                Logger.log(`Using existing destination ID: ${destinationId}`);
                pres = slidesApp.openById(destinationId);
            } catch (e: any) {
                Logger.log(`Error opening destination: ${e.toString()}`);
                throw new Error(`Failed to open destination presentation with ID: ${destinationId}. Details: ${e.message}`);
            }
        } else if (templateId) {
            // 2. Copy Template (Legacy / API mode)
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
            // 3. Create Blank
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
        const diagramGenerator = new GasDiagramSlideGenerator(null);

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
                ...slideModel.rawData // Merge all extra data from JSON
            };

            // Determine Layout
            let slideLayout: GoogleAppsScript.Slides.Layout;
            const layouts = pres.getLayouts();

            // Debug: Log available layouts on first slide only to avoid spam
            if (index === 0) {
                Logger.log('Available Layouts: ' + layouts.map(l => l.getLayoutName()).join(', '));
            }

            // Helper to find layout by name (case-insensitive)
            const findLayout = (name: string) => {
                return layouts.find(l => l.getLayoutName().toUpperCase() === name.toUpperCase());
            };

            const layoutType = (slideModel.layout || 'content').toUpperCase();

            if (layoutType === 'TITLE') {
                slideLayout = findLayout('TITLE') || layouts[0];
            } else if (layoutType === 'SECTION') {
                slideLayout = findLayout('SECTION_HEADER') || findLayout('SECTION ONLY') || findLayout('SECTION TITLE_AND_DESCRIPTION') || layouts[1];
            } else if (layoutType === 'CONTENT' || layoutType === 'AGENDA') {
                slideLayout = findLayout('TITLE_AND_BODY') || findLayout('TITLE_AND_TWO_COLUMNS') || layouts[2];
            } else {
                // For Diagrams and others, usually Title and Body is safest as a canvas
                slideLayout = findLayout('TITLE_AND_BODY') || findLayout('TITLE_ONLY') || layouts[layouts.length - 1];
            }

            // Fallback for safety
            if (!slideLayout) {
                slideLayout = layouts[0];
            }

            Logger.log(`Slide ${index + 1} (${slideModel.layout}): Using Layout '${slideLayout.getLayoutName()}'`);

            const slide = pres.appendSlide(slideLayout);

            // Dispatch to Generators
            const rawType = (slideModel.rawData?.type || '').toLowerCase(); // Original JSON type

            Logger.log(`Dispatching Slide ${index + 1}: LayoutType=${layoutType}, RawType=${rawType}`);

            if (layoutType === 'TITLE') {
                titleGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (layoutType === 'SECTION' || rawType === 'section') {
                sectionGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (
                rawType.includes('timeline') ||
                rawType.includes('process') ||
                rawType.includes('cycle') ||
                rawType.includes('triangle') ||
                rawType.includes('pyramid') ||
                rawType.includes('diagram') ||
                rawType.includes('compare') ||
                rawType.includes('stepup') ||
                rawType.includes('flowchart')
            ) {
                // Use Diagram Generator for visual types
                diagramGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else {
                // Default Content Generator
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
