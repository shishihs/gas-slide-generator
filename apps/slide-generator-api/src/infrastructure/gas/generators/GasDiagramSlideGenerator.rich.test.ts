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
            getRect: vi.fn().mockReturnValue({ left: 0, top: 0, width: 100, height: 100 })
        };
    });

    it('should delegate to DiagramRendererFactory for known types', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = { type: 'cards', items: [] };
        const settings = { primaryColor: '#000' };

        generator.generate(mockSlide, data, mockLayout, 1, settings);

        expect(DiagramRendererFactory.getRenderer).toHaveBeenCalledWith('cards');
        expect(mockRenderer.render).toHaveBeenCalledWith(mockSlide, data, expect.any(Object), settings, mockLayout);
    });

    it('should handle "statsCompare" type', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = { type: 'statsCompare', stats: [] };
        const settings = { primaryColor: '#000' };

        generator.generate(mockSlide, data, mockLayout, 1, settings);

        expect(DiagramRendererFactory.getRenderer).toHaveBeenCalledWith('statscompare'); // lowercased in implementation
        expect(mockRenderer.render).toHaveBeenCalled();
    });

    it('should set title if placeholder exists', () => {
        const mockTitleShape = { getText: vi.fn().mockReturnValue({ setText: vi.fn() }) };
        const mockTitlePlaceholder = { asShape: vi.fn().mockReturnValue(mockTitleShape) };
        mockSlide.getPlaceholder.mockReturnValue(mockTitlePlaceholder);

        const generator = new GasDiagramSlideGenerator(null);
        const data = { type: 'cards', title: 'Test Title' };

        generator.generate(mockSlide, data, mockLayout, 1, {});

        expect(mockTitleShape.getText().setText).toHaveBeenCalledWith('Test Title');
    });
});
