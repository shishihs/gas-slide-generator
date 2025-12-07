import { SlideTitle, SlideContent } from './SlideElement';

export type LayoutType = 'TITLE' | 'AGENDA' | 'CONTENT' | 'SECTION' | 'CONCLUSION';

export class Slide {
    constructor(
        public readonly title: SlideTitle,
        public readonly content: SlideContent,
        public readonly layout: LayoutType = 'CONTENT',
        public readonly subtitle?: string,
        public readonly notes?: string
    ) { }
}
