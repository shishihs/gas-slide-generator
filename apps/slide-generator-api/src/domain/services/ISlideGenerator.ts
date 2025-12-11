
import { LayoutManager } from '../../common/utils/LayoutManager';

export interface ISlideGenerator {
    generate(
        slideId: string,
        data: any,
        layout: LayoutManager,
        pageNum: number,
        settings: any
    ): GoogleAppsScript.Slides.Schema.Request[];
}
