import { Presentation } from '../domain/model/Presentation';
import { Slide } from '../domain/model/Slide';
import { SlideTitle, SlideContent } from '../domain/model/SlideElement';
import { ISlideRepository } from '../domain/repositories/ISlideRepository';

export interface CreatePresentationRequest {
    title: string;
    templateId?: string; // Optional template ID (User provided or Default)
    slides: {
        title: string;
        subtitle?: string;
        content: string[];
        layout?: string; // Matches LayoutType string
        notes?: string;
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

            const slide = new Slide(title, content, layout, slideData.subtitle, slideData.notes);
            presentation.addSlide(slide);
        }

        // Repository now needs to know about templateId. 
        // We should update the interface, but for now let's pass it as a second arg if we update the interface
        return this.slideRepository.createPresentation(presentation, request.templateId);
    }
}
