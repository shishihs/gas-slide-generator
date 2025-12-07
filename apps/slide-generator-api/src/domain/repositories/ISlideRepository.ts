import { Presentation } from '../model/Presentation'; // Fixed relative path
import { Slide } from '../model/Slide'; // Fixed relative path

export interface ISlideRepository {
    createPresentation(presentation: Presentation, templateId?: string): string; // Returns URL
}
