
import { LayoutManager } from '../../common/utils/LayoutManager';

export interface ISlideGenerator {
    generate(
        slide: GoogleAppsScript.Slides.Slide,
        data: any,
        layout: LayoutManager,
        pageNum: number,
        settings: any,
        imageUpdateOption?: string
    ): void;
}
