import { Slide } from './Slide';

export class Presentation {
    private _slides: Slide[] = [];

    constructor(public readonly title: string) { }

    addSlide(slide: Slide) {
        this._slides.push(slide);
    }

    get slides(): Slide[] {
        return [...this._slides];
    }
}
