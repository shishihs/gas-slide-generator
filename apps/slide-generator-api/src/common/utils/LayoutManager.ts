/**
 * LayoutManager - Handles coordinate translation and scaling for slides
 * 
 * Converts pixel-based positions from the theme to point-based positions
 * for the actual slide, accounting for different slide dimensions.
 */

import { SlideTheme, SlidePositions } from '../config/SlideTheme';
import { DEFAULT_THEME } from '../config/DefaultTheme';

export class LayoutManager {
    public scaleX: number;
    public scaleY: number;
    public pageW_pt: number;
    public pageH_pt: number;

    private readonly theme: SlideTheme;
    private readonly pxToPtRatio = 0.75;

    /**
     * Create a new LayoutManager
     * @param pageW_pt - Actual page width in points
     * @param pageH_pt - Actual page height in points
     * @param theme - Theme configuration (defaults to DEFAULT_THEME)
     */
    constructor(pageW_pt: number, pageH_pt: number, theme: SlideTheme = DEFAULT_THEME) {
        this.theme = theme;
        this.pageW_pt = pageW_pt;
        this.pageH_pt = pageH_pt;

        const baseW_pt = this.pxToPt(theme.basePx.width);
        const baseH_pt = this.pxToPt(theme.basePx.height);
        this.scaleX = pageW_pt / baseW_pt;
        this.scaleY = pageH_pt / baseH_pt;
    }

    /**
     * Convert pixels to points
     */
    public pxToPt(px: number): number {
        return px * this.pxToPtRatio;
    }

    /**
     * Get position from a dot-notation path (e.g., "contentSlide.title")
     */
    private getPositionFromPath(path: string): any {
        return path.split('.').reduce(
            (obj: any, key) => obj && obj[key],
            this.theme.positions
        );
    }

    /**
     * Get a rectangle specification converted to points and scaled
     * @param spec - Either a dot-notation path string or a position object
     */
    public getRect(spec: string | any) {
        const pos = typeof spec === 'string' ? this.getPositionFromPath(spec) : spec;
        if (!pos) return { left: 0, top: 0, width: 0, height: 0 };

        let left_px = pos.left;
        if (pos.right !== undefined && pos.left === undefined) {
            left_px = this.theme.basePx.width - pos.right - pos.width;
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

    /**
     * Get the current theme
     */
    public getTheme(): SlideTheme {
        return this.theme;
    }
}
