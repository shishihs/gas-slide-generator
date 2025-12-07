import { PresentationApplicationService, CreatePresentationRequest } from '../application/PresentationApplicationService';
import { GasSlideRepository } from '../infrastructure/gas/GasSlideRepository';

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
function doPost(e: GoogleAppsScript.Events.DoPost) {
    try {
        const postData = JSON.parse(e.postData.contents);

        // Extract JSON payload (handling wrapped structure if needed)
        const requestData: CreatePresentationRequest = postData.json || postData;

        // Dependency Injection Composition Root
        const repository = new GasSlideRepository();
        const service = new PresentationApplicationService(repository);

        const slideUrl = service.createPresentation(requestData);

        return createJsonResponse({
            success: true,
            url: slideUrl
        });

    } catch (error: any) {
        Logger.log('API Error: ' + error.toString());
        return createJsonResponse({
            success: false,
            error: error.message || error.toString()
        });
    }
}

function createJsonResponse(data: SlideGenerationResponse) {
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}
