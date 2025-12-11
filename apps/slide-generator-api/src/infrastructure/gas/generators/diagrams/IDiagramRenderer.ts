import { LayoutManager } from '../../../../common/utils/LayoutManager';

export interface IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void;
}
