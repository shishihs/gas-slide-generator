
import { GasDiagramSlideGenerator } from './GasDiagramSlideGenerator';
import { DiagramRendererFactory } from './diagrams/DiagramRendererFactory';
import { LayoutManager } from '../../../common/utils/LayoutManager';

// Mock Dependencies
jest.mock('../../../common/utils/SlideUtils', () => ({
    setStyledText: jest.fn(),
    offsetRect: jest.fn(),
    addFooter: jest.fn(),
    drawArrowBetweenRects: jest.fn(),
    setBoldTextSize: jest.fn(),
    insertImageFromUrlOrFileId: jest.fn()
}));

jest.mock('../../../common/utils/ColorUtils', () => ({
    generateProcessColors: jest.fn(),
    generateTimelineCardColors: jest.fn(),
    generatePyramidColors: jest.fn(),
    generateCompareColors: jest.fn(),
    generateTintedGray: jest.fn()
}));

jest.mock('./diagrams/DiagramRendererFactory');

describe('GasDiagramSlideGenerator', () => {
    let mockRenderer;
    let mockSlide;
    let mockLayout;

    beforeEach(() => {
        jest.clearAllMocks();

        // Mock global GAS objects
        (global as any).SlidesApp = {
            ShapeType: { RECTANGLE: 'RECTANGLE' } as any,
            PlaceholderType: { TITLE: 'TITLE', CENTERED_TITLE: 'CENTERED_TITLE', BODY: 'BODY', OBJECT: 'OBJECT', PICTURE: 'PICTURE' } as any,
            ParagraphAlignment: { START: 'START' } as any,
            ContentAlignment: { MIDDLE: 'MIDDLE' } as any
        } as any;
        (global as any).Logger = { log: jest.fn() } as any;

        // Setup Factory Mock
        mockRenderer = { render: jest.fn() };
        (DiagramRendererFactory.getRenderer as jest.Mock).mockReturnValue(mockRenderer);

        // Setup Slide Mock
        mockSlide = {
            getPlaceholder: jest.fn().mockReturnValue(null),
            getPlaceholders: jest.fn().mockReturnValue([]),
            getPageElements: jest.fn().mockReturnValue([]),
            group: jest.fn(),
            insertShape: jest.fn()
        };

        // Setup Layout Mock
        mockLayout = {
            getRect: jest.fn().mockReturnValue({ left: 0, top: 0, width: 100, height: 100 })
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
        const mockTitleShape = { getText: jest.fn().mockReturnValue({ setText: jest.fn() }) };
        const mockTitlePlaceholder = { asShape: jest.fn().mockReturnValue(mockTitleShape) };
        mockSlide.getPlaceholder.mockReturnValue(mockTitlePlaceholder);

        const generator = new GasDiagramSlideGenerator(null);
        const data = { type: 'cards', title: 'Test Title' };

        generator.generate(mockSlide, data, mockLayout, 1, {});

        expect(mockTitleShape.getText().setText).toHaveBeenCalledWith('Test Title');
    });
});
