import { vi, describe, it, expect, beforeEach, Mock } from 'vitest';
import { GasDiagramSlideGenerator } from './GasDiagramSlideGenerator';
import { DiagramRendererFactory } from './diagrams/DiagramRendererFactory';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { setupGasGlobals } from '../../../test/gas-mocks';

// Mock Dependencies
vi.mock('../../../common/utils/SlideUtils', () => ({
    setStyledText: vi.fn(),
    offsetRect: vi.fn(),
    addFooter: vi.fn(),
    drawArrowBetweenRects: vi.fn(),
    setBoldTextSize: vi.fn(),
    insertImageFromUrlOrFileId: vi.fn()
}));

vi.mock('../../../common/utils/ColorUtils', () => ({
    generateProcessColors: vi.fn(),
    generateTimelineCardColors: vi.fn(),
    generatePyramidColors: vi.fn(),
    generateCompareColors: vi.fn(),
    generateTintedGray: vi.fn()
}));

vi.mock('./diagrams/DiagramRendererFactory');

describe('GasDiagramSlideGenerator', () => {
    let mockRenderer;
    let mockSlide;
    let mockLayout;

    beforeEach(() => {
        vi.clearAllMocks();

        // Setup Type-Safe Global Mocks
        setupGasGlobals();

        // Setup Factory Mock
        mockRenderer = { render: vi.fn() };
        (DiagramRendererFactory.getRenderer as Mock).mockReturnValue(mockRenderer);

        // Setup Slide Mock
        mockSlide = {
            getPlaceholder: vi.fn().mockReturnValue(null),
            getPlaceholders: vi.fn().mockReturnValue([]),
            getPageElements: vi.fn().mockReturnValue([]),
            group: vi.fn(),
            insertShape: vi.fn()
        };

        // Setup Layout Mock
        mockLayout = {
            getRect: vi.fn().mockReturnValue({ left: 0, top: 0, width: 100, height: 100 }),
            getTheme: vi.fn().mockReturnValue({
                fonts: { family: 'Arial', sizes: { title: 40, body: 20 } },
                colors: { primary: '#000000', neutralGray: '#666666' }
            })
        };
    });

    it('should delegate to DiagramRendererFactory for known types', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = { type: 'cards', items: [] };
        const settings = { primaryColor: '#000' };

        // Mock renderer to return empty requests
        mockRenderer.render.mockReturnValue([]);

        const result = generator.generate('slide-id', data, mockLayout as any, 1, settings);

        expect(DiagramRendererFactory.getRenderer).toHaveBeenCalledWith('cards');
        expect(mockRenderer.render).toHaveBeenCalledWith('slide-id', data, expect.any(Object), settings, mockLayout);
        expect(Array.isArray(result)).toBe(true);
    });

    it('should handle "statsCompare" type', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = { type: 'statsCompare', stats: [] };
        const settings = { primaryColor: '#000' };

        mockRenderer.render.mockReturnValue([]);

        const result = generator.generate('slide-id', data, mockLayout as any, 1, settings);

        expect(DiagramRendererFactory.getRenderer).toHaveBeenCalledWith('statscompare');
        expect(mockRenderer.render).toHaveBeenCalled();
        expect(Array.isArray(result)).toBe(true);
    });

    it('should generate title requests if title is present', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = { type: 'cards', title: 'Test Title' };

        mockRenderer.render.mockReturnValue([]);

        const result = generator.generate('slide-id', data, mockLayout as any, 1, {});

        // Check for insertText request
        const textReq = result.find(r => r.insertText && r.insertText.text === 'Test Title');
        expect(textReq).toBeDefined();
    });
});
