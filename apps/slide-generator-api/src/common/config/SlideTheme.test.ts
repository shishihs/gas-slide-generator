import { SlideTheme, mergeTheme, PartialTheme } from './SlideTheme';
import { DEFAULT_THEME } from './DefaultTheme';

describe('SlideTheme', () => {
    describe('DEFAULT_THEME', () => {
        it('should have all required properties', () => {
            expect(DEFAULT_THEME.basePx).toBeDefined();
            expect(DEFAULT_THEME.basePx.width).toBe(960);
            expect(DEFAULT_THEME.basePx.height).toBe(540);

            expect(DEFAULT_THEME.fonts).toBeDefined();
            expect(DEFAULT_THEME.fonts.family).toBe('Noto Sans JP');
            expect(DEFAULT_THEME.fonts.sizes.body).toBe(16);

            expect(DEFAULT_THEME.colors).toBeDefined();
            expect(DEFAULT_THEME.colors.primary).toBe('#4A6C42');
            expect(DEFAULT_THEME.colors.textPrimary).toBe('#212121');

            expect(DEFAULT_THEME.diagram).toBeDefined();
            expect(DEFAULT_THEME.diagram.laneGapPx).toBe(30);

            expect(DEFAULT_THEME.positions).toBeDefined();
            expect(DEFAULT_THEME.positions.contentSlide).toBeDefined();
        });
    });

    describe('mergeTheme', () => {
        it('should return base theme when partial is empty', () => {
            const result = mergeTheme(DEFAULT_THEME, {});
            expect(result).toEqual(DEFAULT_THEME);
        });

        it('should override colors.primary', () => {
            const partial: PartialTheme = {
                colors: { primary: '#FF5722' }
            };
            const result = mergeTheme(DEFAULT_THEME, partial);

            expect(result.colors.primary).toBe('#FF5722');
            // Other colors should remain unchanged
            expect(result.colors.textPrimary).toBe(DEFAULT_THEME.colors.textPrimary);
        });

        it('should override fonts.family', () => {
            const partial: PartialTheme = {
                fonts: { family: 'Roboto' }
            };
            const result = mergeTheme(DEFAULT_THEME, partial);

            expect(result.fonts.family).toBe('Roboto');
            // Font sizes should remain unchanged
            expect(result.fonts.sizes.body).toBe(DEFAULT_THEME.fonts.sizes.body);
        });

        it('should override fonts.sizes (full object)', () => {
            const partial: PartialTheme = {
                fonts: {
                    sizes: {
                        ...DEFAULT_THEME.fonts.sizes,
                        body: 16
                    }
                }
            };
            const result = mergeTheme(DEFAULT_THEME, partial);

            expect(result.fonts.sizes.body).toBe(16);
            expect(result.fonts.sizes.title).toBe(DEFAULT_THEME.fonts.sizes.title);
        });

        it('should override footerText', () => {
            const partial: PartialTheme = {
                footerText: 'Custom Footer'
            };
            const result = mergeTheme(DEFAULT_THEME, partial);

            expect(result.footerText).toBe('Custom Footer');
        });

        it('should override multiple properties at once', () => {
            const partial: PartialTheme = {
                colors: { primary: '#3498db' },
                fonts: { family: 'Arial' },
                footerText: 'My Company'
            };
            const result = mergeTheme(DEFAULT_THEME, partial);

            expect(result.colors.primary).toBe('#3498db');
            expect(result.fonts.family).toBe('Arial');
            expect(result.footerText).toBe('My Company');
        });
    });
});
