import { SlideTitle, SlideContent } from './SlideElement';

export class Slide {
    constructor(
        public readonly title: SlideTitle,
        public readonly content: SlideContent,
        public readonly notes?: string
    ) { }
}
