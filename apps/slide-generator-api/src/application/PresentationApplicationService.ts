import { Presentation } from '../domain/model/Presentation';
import { Slide } from '../domain/model/Slide';
import { SlideTitle, SlideContent } from '../domain/model/SlideElement';
import { ISlideRepository } from '../domain/repositories/ISlideRepository';

export interface CreatePresentationRequest {
    title: string;
    slides: {
        title: string;
        content: string[];
        notes?: string;
    }[];
}

export class PresentationApplicationService {
    constructor(private readonly slideRepository: ISlideRepository) { }

    createPresentation(request: CreatePresentationRequest): string {
        const presentation = new Presentation(request.title);

        for (const slideData of request.slides) {
            const title = new SlideTitle(slideData.title);
            const content = new SlideContent(slideData.content);
            const slide = new Slide(title, content, slideData.notes);
            presentation.addSlide(slide);
        }

        return this.slideRepository.createPresentation(presentation);
    }
}
