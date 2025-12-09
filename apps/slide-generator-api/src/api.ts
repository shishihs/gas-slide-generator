import { PresentationApplicationService, CreatePresentationRequest } from './application/PresentationApplicationService';
import { GasSlideRepository } from './infrastructure/gas/GasSlideRepository';

/**
 * Slide Generator API
 * 
 * Receives JSON from Document Add-on and generates Google Slides.
 */

interface SlideGenerationResponse {
    url?: string;
    error?: string;
    success: boolean;
}

/**
 * Web App Entry Point: POST
 */
// Explicitly expose functions to global scope for GAS
(global as any).doPost = doPost;
(global as any).doGet = doGet;
(global as any).generateSlides = generateSlides;

/**
 * Library Entry Point: generateSlides
 * 
 * Callable directly from other scripts when added as a library.
 * 
 * @param {any} data - The JSON payload (Presentation Data)
 * @returns {SlideGenerationResponse}
 */
function generateSlides(data: any): SlideGenerationResponse {
    try {
        Logger.log('Library Call: generateSlides with data: ' + JSON.stringify(data));

        const request: CreatePresentationRequest = {
            title: data.title || 'Untitled Presentation',
            templateId: data.templateId, // Optional ID for template
            destinationId: data.destinationId, // Optional ID for existing destination
            slides: data.slides.map((s: any) => ({
                ...s, // Spread all valid properties from source
                title: s.title,
                subtitle: s.subtitle || s.subhead, // Handle alias if inconsistent
                content: s.content || s.points || [], // Partial alias support
                layout: s.layout || s.type, // Map type to layout
                notes: s.notes
            }))
        };

        // Dependency Injection Composition Root
        const repository = new GasSlideRepository();
        const service = new PresentationApplicationService(repository);

        const slideUrl = service.createPresentation(request);

        return {
            success: true,
            url: slideUrl
        };

    } catch (error: any) {
        Logger.log('Library Error: ' + error.toString());
        return {
            success: false,
            error: error.message || error.toString()
        };
    }
}

function doGet(e: GoogleAppsScript.Events.DoGet) {
    return ContentService.createTextOutput("Slide Generator API is running. Authorization successful.");
}

function doPost(e: GoogleAppsScript.Events.DoPost) {
    try {
        const postData = JSON.parse(e.postData.contents);

        // Handle Test Connection
        if (postData.action === 'test') {
            return createJsonResponse({
                success: true,
                message: 'POST Connection Successful',
                received: postData
            });
        }

        // Extract JSON payload (handling wrapped structure if needed)
        const data = postData.json || postData;
        Logger.log('Incoming Request Data: ' + JSON.stringify(data));

        // Wrap logic by calling the core logic function
        const result = generateSlides(data);

        return createJsonResponse(result);

    } catch (error: any) {
        Logger.log('API Error: ' + error.toString());
        return createJsonResponse({
            success: false,
            error: error.message || error.toString()
        });
    }
}

(global as any).doPost = doPost;

function createJsonResponse(data: SlideGenerationResponse) {
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}
