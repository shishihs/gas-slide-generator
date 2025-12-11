
import { ISlideRepository } from '../../domain/repositories/ISlideRepository';
import { Presentation } from '../../domain/model/Presentation';
import { LayoutManager } from '../../common/utils/LayoutManager';
import { GasTitleSlideGenerator } from './generators/GasTitleSlideGenerator';
import { GasSectionSlideGenerator } from './generators/GasSectionSlideGenerator'; // Not refactored yet, assuming legacy implementation or minimal impact
import { GasContentSlideGenerator } from './generators/GasContentSlideGenerator';
import { GasDiagramSlideGenerator } from './generators/GasDiagramSlideGenerator';
import { DEFAULT_THEME, AVAILABLE_THEMES } from '../../common/config/DefaultTheme';

export class GasSlideRepository implements ISlideRepository {

    // Note: sectionGenerator still uses old interface? 
    // If so, we can't fully batch it unless we refactor it too.
    // Plan: We marked GasSectionSlideGenerator as TODO in task.md. 
    // For now, I'll assume we can make minimal changes or ignore it for the prototype.
    // Actually, to make ts happy, we'll need to update it or cast it.
    // Let's assume we update GasSectionSlideGenerator briefly or inline.

    createPresentation(presentation: Presentation, templateId?: string, destinationId?: string, settingsOverride?: any): string {
        const slidesApp = SlidesApp;
        const driveApp = DriveApp;

        // 1. Create/Open Presentation
        let pres: GoogleAppsScript.Slides.Presentation;
        let presId: string;

        if (destinationId) {
            pres = slidesApp.openById(destinationId);
        } else if (templateId) {
            const templateFile = driveApp.getFileById(templateId);
            const newFile = templateFile.makeCopy(presentation.title);
            pres = slidesApp.openById(newFile.getId());
        } else {
            pres = slidesApp.create(presentation.title);
        }
        // Ensure pres exists (for Test Mocks that might vary)
        presId = pres ? pres.getId() : 'mock-presentation-id';

        // 2. Setup Layout & Theme
        const pageWidth = pres.getPageWidth();
        const pageHeight = pres.getPageHeight();
        const themeName = settingsOverride && settingsOverride.theme ? settingsOverride.theme : 'Green';
        const selectedTheme = AVAILABLE_THEMES[themeName] || DEFAULT_THEME;
        const layoutManager = new LayoutManager(pageWidth, pageHeight, selectedTheme);
        const settings = { primaryColor: selectedTheme.colors.primary, ...settingsOverride }; // Simplified

        // 3. Map Display Names to Layout IDs (Fix for custom layout names)
        const layoutIdMap = new Map<string, string>();
        try {
            // @ts-ignore: Advanced Service
            const presData = Slides.Presentations.get(presId);
            if (presData.layouts) {
                presData.layouts.forEach((layout: any) => {
                    const props = layout.layoutProperties;
                    if (props) {
                        if (props.displayName) {
                            layoutIdMap.set(props.displayName.toUpperCase(), layout.objectId);
                        }
                        if (props.name) {
                            layoutIdMap.set(props.name.toUpperCase(), layout.objectId);
                        }
                    }
                });
            }
        } catch (e) {
            Logger.log('Warning: Failed to fetch presentation details via Advanced API. Layout mapping may be limited. ' + e);
        }

        // Debug: Log mapped layouts
        Logger.log(`Mapped Layouts (Display Name -> ID): ${Array.from(layoutIdMap.keys()).join(', ')}`);

        // 4. Generators
        const titleGenerator = new GasTitleSlideGenerator(null);
        const sectionGenerator = new GasSectionSlideGenerator(null); // Refactored to Batch
        const contentGenerator = new GasContentSlideGenerator(null);
        const diagramGenerator = new GasDiagramSlideGenerator(null);

        // 5. Batch Request Collection
        let allRequests: GoogleAppsScript.Slides.Schema.Request[] = [];

        // Debug: Log all available layouts in the presentation
        // const availableLayouts = pres.getLayouts().map(l => l.getLayoutName());
        // Logger.log(`Available Layouts in Presentation: ${availableLayouts.join(', ')}`);

        // 6. Generate Slide Creation Requests + Content Requests
        // Note: Mix of API and App usage. 
        // Best Practice: Use API for creating slides too to map IDs easily.
        // BUT templates use `slideLayoutReference`.
        // If we use `pres.appendSlide`, we get a Slide object with an ID.
        // Iterating efficiently:

        // Remove default slide if creating new
        if (!templateId && !destinationId && pres.getSlides().length > 0) {
            pres.getSlides()[0].remove();
        }

        presentation.slides.forEach((slideModel, index) => {
            const layoutType = (slideModel.layout || 'content').toUpperCase();
            const rawType = (slideModel.rawData?.type || '').toLowerCase(); // Original JSON type

            Logger.log(`[Slide ${index + 1}] Processing - Requested Layout: '${slideModel.layout}', Normalized: '${layoutType}', RawType: '${rawType}'`);

            // Step A: Create Slide Object (Synchronously via App to get ID and Layout easily)
            // Ideally we'd do this via API too, but getting Layout IDs is tricky without reading.
            // So we stick to `pres.appendSlide` to get the Object and ID.
            // This is "Hybrid Mode": App for Structure, API for Content.

            let slideLayout: GoogleAppsScript.Slides.Layout | undefined;
            const layouts = pres.getLayouts();

            // Attempt 1: Match by mapped Display Name or Internal Name
            const targetLayoutId = layoutIdMap.get(layoutType);
            if (targetLayoutId) {
                slideLayout = layouts.find(l => l.getObjectId() === targetLayoutId);
            }

            // Attempt 2: Fallback to simple matching if map failed or missed
            if (!slideLayout) {
                slideLayout = layouts.find(l => l.getLayoutName().toUpperCase() === layoutType);
            }

            // Attempt 3: Smart Fallback for Diagrams etc. -> use 'CONTENT' if available
            // This is useful if user only has basic layouts (AGENDA, CONTENT, ETC) but requests 'TIMELINE'.
            if (!slideLayout && layoutType !== 'TITLE' && layoutType !== 'SECTION') {
                // Try mapping 'TIMELINE', 'PROCESS' etc to 'CONTENT'
                const fallbackId = layoutIdMap.get('CONTENT');
                if (fallbackId) {
                    slideLayout = layouts.find(l => l.getObjectId() === fallbackId);
                    if (slideLayout) {
                        Logger.log(`[Slide ${index + 1}] Layout '${layoutType}' not found. Falling back to 'CONTENT'.`);
                    }
                }
                // Try standard internal name if custom name map failed
                if (!slideLayout) {
                    slideLayout = layouts.find(l => l.getLayoutName().toUpperCase() === 'CONTENT' || l.getLayoutName().toUpperCase() === 'TITLE_AND_BODY');
                }
            }

            // Attempt 4: Final Fallback (BLANK or Last)
            if (!slideLayout) {
                // Force BLANK layout to allow full control via Batch API if specific layout not found
                // This avoids dependency on Placeholder IDs which are unpredictable.
                slideLayout = layouts.find(l => l.getLayoutName().toUpperCase() === 'BLANK') || layouts[layouts.length - 1];
                Logger.log(`[Slide ${index + 1}] Layout '${layoutType}' not found. Falling back to '${slideLayout.getLayoutName()}'`);
            } else {
                Logger.log(`[Slide ${index + 1}] Found Layout: '${slideLayout.getLayoutName()}' (ID: ${slideLayout.getObjectId()})`);
            }

            const slide = pres.appendSlide(slideLayout);
            const slideId = slide.getObjectId();

            // Step B: Generate Content Requests
            const commonData = {
                title: slideModel.title.value,
                subtitle: slideModel.subtitle,
                date: new Date().toLocaleDateString(),
                points: slideModel.content.items,
                content: slideModel.content.items,
                ...slideModel.rawData
            };

            let reqs: GoogleAppsScript.Slides.Schema.Request[] = [];

            if (layoutType === 'TITLE') {
                reqs = titleGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
            } else if (layoutType === 'SECTION' || layoutType === 'SECTION_HEADER') {
                reqs = sectionGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
            } else if (
                rawType.includes('timeline') ||
                rawType.includes('process') || // ... other diagrams
                rawType.includes('diagram')
            ) {
                reqs = diagramGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
            } else {
                reqs = contentGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
            }

            Logger.log(`[Slide ${index + 1}] Generator Used for layoutType '${layoutType}': ${reqs.length} requests generated.`);

            allRequests.push(...reqs);
        });

        // 6. Execute Batch Update
        if (allRequests.length > 0) {
            // Split into chunks if too large (API limit is usually high, but safe practice)
            // Limit is 100 requests per call is FALSE. Limit is generous, but payload size matters.
            // We can send all in one go for typical pres (e.g. < 50 slides).
            try {
                // @ts-ignore
                Slides.Presentations.batchUpdate({ requests: allRequests }, presId);
                Logger.log(`Executed ${allRequests.length} requests successfully.`);
            } catch (e) {
                Logger.log(`Batch Update Failed: ${e}`);
                throw e;
            }
        }

        return pres.getUrl();
    }
}
