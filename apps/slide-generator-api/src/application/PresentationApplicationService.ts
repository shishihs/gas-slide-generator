import { Presentation } from '../domain/model/Presentation';
import { Slide } from '../domain/model/Slide';
import { SlideTitle, SlideContent } from '../domain/model/SlideElement';
import { ISlideRepository } from '../domain/repositories/ISlideRepository';

export interface CreatePresentationRequest {
    title: string;
    templateId?: string; // Optional template ID (for copying)
    destinationId?: string; // Optional destination ID (pre-copied presentation)
    slides: {
        title: string;
        subtitle?: string;
        content: string[];
        layout?: string;
        notes?: string;
        [key: string]: any; // Allow arbitrary extra properties
    }[];
}

export class PresentationApplicationService {
    constructor(private readonly slideRepository: ISlideRepository) { }

    createPresentation(request: CreatePresentationRequest): string {
        const presentation = new Presentation(request.title);

        // Pass template ID to presentation if needed, or handle in repository
        // For now, we likely need to pass it to the repository method or store in Presentation entity
        // Let's assume Repository handles it via a second argument or Presentation has metadata

        for (const slideData of request.slides) {
            const title = new SlideTitle(slideData.title);
            const content = new SlideContent(slideData.content);
            // Cast string to LayoutType, defaulting to 'CONTENT' if undefined
            const layout = (slideData.layout as any) || 'CONTENT';

            // Pass the entire slideData as rawData
            const slide = new Slide(title, content, layout, slideData.subtitle, slideData.notes, slideData);
            presentation.addSlide(slide);
        }

        // Repository now needs to know about templateId. 
        // We should update the interface, but for now let's pass it as a second arg if we update the interface
        return this.slideRepository.createPresentation(presentation, request.templateId, request.destinationId);
    }
}
