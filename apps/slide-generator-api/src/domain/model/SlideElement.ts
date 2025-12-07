export class SlideTitle {
    constructor(public readonly value: string) {
        if (!value) {
            // Allow empty titles for now, or throw validation error
            // throw new Error("Slide title cannot be empty"); 
        }
    }
}

export class SlideContent {
    constructor(public readonly items: string[]) { }
}

export class Slide {
    constructor(
        public readonly title: SlideTitle,
        public readonly content: SlideContent,
        public readonly notes?: string
    ) { }
}
