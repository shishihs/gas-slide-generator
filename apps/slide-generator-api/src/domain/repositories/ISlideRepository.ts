import { Presentation } from './Presentation';
import { Slide } from './Slide';

export interface ISlideRepository {
    createPresentation(presentation: Presentation): string; // Returns URL
}
