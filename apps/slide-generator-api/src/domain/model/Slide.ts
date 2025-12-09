import { SlideTitle, SlideContent } from './SlideElement';

export type LayoutType = string;

export class Slide {
    constructor(
        public readonly title: SlideTitle,
        public readonly content: SlideContent,
        public readonly layout: LayoutType = 'CONTENT',
        public readonly subtitle?: string,
        public readonly notes?: string,
        public readonly rawData?: any // Store extra data from JSON
    ) { }
}
