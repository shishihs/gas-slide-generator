import { LayoutManager } from '../../../../common/utils/LayoutManager';

export interface IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[];
}
