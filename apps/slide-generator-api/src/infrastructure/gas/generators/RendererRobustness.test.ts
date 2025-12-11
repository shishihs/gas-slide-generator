import { vi, describe, it, expect, beforeEach } from 'vitest';
import { LanesDiagramRenderer } from './diagrams/LanesDiagramRenderer';
import { BarCompareDiagramRenderer } from './diagrams/BarCompareDiagramRenderer';
import { QuoteDiagramRenderer } from './diagrams/QuoteDiagramRenderer';
import { ProcessDiagramRenderer } from './diagrams/ProcessDiagramRenderer';
import { setupGasGlobals } from '../../../test/gas-mocks';

// Mock Dependencies
vi.mock('../../../common/utils/SlideUtils', () => ({
    setStyledText: vi.fn(),
    offsetRect: vi.fn(),
    addFooter: vi.fn(),
    drawArrowBetweenRects: vi.fn(),
    setBoldTextSize: vi.fn()
}));

describe('Renderer Robustness', () => {
    let mockLayout: any;
    const area = { left: 0, top: 0, width: 960, height: 540 };

    beforeEach(() => {
        vi.clearAllMocks();
        setupGasGlobals();

        mockLayout = {
            getRect: vi.fn().mockReturnValue({ left: 0, top: 0, width: 960, height: 540 }),
            pxToPt: vi.fn((px) => px * 0.75),
            getTheme: vi.fn().mockReturnValue({
                colors: {
                    primary: '#4285F4',
                    textPrimary: '#000000',
                    neutralGray: '#888888',
                    backgroundGray: '#F1F3F4',
                    laneBorder: '#E0E0E0',
                    cardBorder: '#CCCCCC',
                    ghostGray: '#F5F5F5',
                    faintGray: '#F8F9FA'
                },
                fonts: {
                    family: 'Roboto',
                    sizes: { title: 32, body: 14, laneTitle: 18 }
                },
                diagram: {
                    laneGapPx: 10, lanePadPx: 10, laneTitleHeightPx: 40,
                    cardGapPx: 10, cardMinHeightPx: 60, cardMaxHeightPx: 120,
                    arrowHeightPx: 20, arrowGapPx: 5
                }
            })
        };
    });

    describe('LanesDiagramRenderer', () => {
        it('should handle null primaryColor in settings', () => {
            const renderer = new LanesDiagramRenderer();
            const data = {
                lanes: [
                    { title: 'Lane 1', items: ['Item 1', 'Item 2'] },
                    { title: 'Lane 2', items: ['Item 3'] }
                ]
            };
            // @ts-ignore
            const settings = { primaryColor: null };

            // Should not throw
            const reqs = renderer.render('slide-id', data, area, settings, mockLayout);

            expect(reqs.length).toBeGreaterThan(0);
        });

        it('should handle undefined primaryColor in settings', () => {
            const renderer = new LanesDiagramRenderer();
            const data = { lanes: [{ title: 'L1', items: [] }] };
            const settings = {};

            const reqs = renderer.render('slide-id', data, area, settings, mockLayout);
            expect(reqs.length).toBeGreaterThan(0);
        });
    });

    describe('BarCompareDiagramRenderer', () => {
        it('should handle null primaryColor', () => {
            const renderer = new BarCompareDiagramRenderer();
            const data = {
                stats: [{ label: 'A', leftValue: '10', rightValue: '20' }]
            };
            // @ts-ignore
            const settings = { primaryColor: null };
            const reqs = renderer.render('slide-id', data, area, settings, mockLayout);
            expect(reqs.length).toBeGreaterThan(0);
        });
    });

    describe('QuoteDiagramRenderer', () => {
        it('should handle null primaryColor', () => {
            const renderer = new QuoteDiagramRenderer();
            const data = { quote: 'Hello', author: 'Me' };
            // @ts-ignore
            const settings = { primaryColor: null };
            const reqs = renderer.render('slide-id', data, area, settings, mockLayout);
            expect(reqs.length).toBeGreaterThan(0);
        });
    });

    describe('ProcessDiagramRenderer', () => {
        it('should handle invalid hex values', () => {
            const renderer = new ProcessDiagramRenderer();
            const data = { steps: ['Step 1'] };
            // @ts-ignore
            const settings = { primaryColor: 'invalid-hex' };
            const reqs = renderer.render('slide-id', data, area, settings, mockLayout);
            expect(reqs.length).toBeGreaterThan(0);
        });
    });
});
