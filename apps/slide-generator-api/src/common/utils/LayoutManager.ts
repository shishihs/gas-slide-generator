
import { CONFIG } from '../config/SlideConfig';

export class LayoutManager {
    public scaleX: number;
    public scaleY: number;
    public pageW_pt: number;
    public pageH_pt: number;
    private pxToPtFn: (px: number) => number;

    constructor(pageW_pt: number, pageH_pt: number) {
        this.pxToPtFn = (px) => px * 0.75;
        const baseW_pt = this.pxToPt(CONFIG.BASE_PX.W);
        const baseH_pt = this.pxToPt(CONFIG.BASE_PX.H);
        this.scaleX = pageW_pt / baseW_pt;
        this.scaleY = pageH_pt / baseH_pt;
        this.pageW_pt = pageW_pt;
        this.pageH_pt = pageH_pt;
    }

    public pxToPt(px: number): number {
        return this.pxToPtFn(px);
    }

    private getPositionFromPath(path: string) {
        return path.split('.').reduce((obj: any, key) => obj && obj[key], CONFIG.POS_PX);
    }

    public getRect(spec: string | any) {
        const pos = typeof spec === 'string' ? this.getPositionFromPath(spec) : spec;
        if (!pos) return { left: 0, top: 0, width: 0, height: 0 };

        let left_px = pos.left;
        if (pos.right !== undefined && pos.left === undefined) {
            left_px = CONFIG.BASE_PX.W - pos.right - pos.width;
        }
        if (left_px === undefined && pos.right === undefined) {
            left_px = 0;
        }

        return {
            left: left_px !== undefined ? this.pxToPt(left_px) * this.scaleX : 0,
            top: pos.top !== undefined ? this.pxToPt(pos.top) * this.scaleY : 0,
            width: pos.width !== undefined ? this.pxToPt(pos.width) * this.scaleX : 0,
            height: pos.height !== undefined ? this.pxToPt(pos.height) * this.scaleY : 0,
        };
    }
}
