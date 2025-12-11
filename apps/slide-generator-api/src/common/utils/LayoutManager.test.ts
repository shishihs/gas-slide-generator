import { LayoutManager } from './LayoutManager';
import { DEFAULT_THEME } from '../config/DefaultTheme';
import { SlideTheme, mergeTheme } from '../config/SlideTheme';

describe('LayoutManager', () => {
    const pageW = 720; // pt
    const pageH = 405; // pt

    describe('constructor', () => {
        it('should use DEFAULT_THEME when no theme is provided', () => {
            const layout = new LayoutManager(pageW, pageH);
            const theme = layout.getTheme();

            expect(theme).toBe(DEFAULT_THEME);
        });

        it('should use custom theme when provided', () => {
            const customTheme = mergeTheme(DEFAULT_THEME, {
                colors: { primary: '#FF0000' }
            });

            const layout = new LayoutManager(pageW, pageH, customTheme);
            const theme = layout.getTheme();

            expect(theme.colors.primary).toBe('#FF0000');
        });
    });

    describe('pxToPt', () => {
        it('should convert pixels to points based on page dimensions', () => {
            const layout = new LayoutManager(pageW, pageH);

            // 960px base → 720pt page = 0.75 ratio
            const result = layout.pxToPt(100);
            expect(result).toBe(75); // 100 * 0.75
        });
    });

    describe('getRect', () => {
        it('should return scaled rectangle for contentSlide.body', () => {
            const layout = new LayoutManager(pageW, pageH);
            const rect = layout.getRect('contentSlide.body');

            expect(rect).toBeDefined();
            expect(rect.left).toBeDefined();
            expect(rect.top).toBeDefined();
            expect(rect.width).toBeDefined();
            expect(rect.height).toBeDefined();
        });

        it('should return empty rect for non-existent path', () => {
            const layout = new LayoutManager(pageW, pageH);
            const rect = layout.getRect('nonExistent.path');

            // LayoutManager returns empty rect (0,0,0,0) for unknown paths
            expect(rect.width).toBe(0);
            expect(rect.height).toBe(0);
        });

        it('should return scaled rectangle for titleSlide.title', () => {
            const layout = new LayoutManager(pageW, pageH);
            const rect = layout.getRect('titleSlide.title');

            expect(rect).toBeDefined();
            // Verify scaling: original is 60px left at 960px base
            // At 720pt page with 0.75 ratio: 60 * 0.75 = 45
            expect(rect.left).toBeCloseTo(45, 1);
        });

        it('should handle positions with right property', () => {
            const layout = new LayoutManager(pageW, pageH);
            const rect = layout.getRect('contentSlide.headerLogo');

            expect(rect).toBeDefined();
            // right: 20px → scaled and converted to left position
            expect(rect.left).toBeDefined();
        });
    });

    describe('getTheme', () => {
        it('should return the injected theme', () => {
            const customTheme = mergeTheme(DEFAULT_THEME, {
                footerText: 'Test Footer'
            });

            const layout = new LayoutManager(pageW, pageH, customTheme);

            expect(layout.getTheme().footerText).toBe('Test Footer');
        });
    });
});
