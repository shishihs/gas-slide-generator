
import { GasDiagramSlideGenerator } from './GasDiagramSlideGenerator';

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
    generateProcessColors: jest.fn().mockReturnValue(['#ff0000', '#00ff00', '#0000ff']),
    generateTimelineCardColors: jest.fn().mockReturnValue(['#ff0000', '#00ff00', '#0000ff']),
    generatePyramidColors: jest.fn().mockReturnValue(['#ff0000', '#00ff00', '#0000ff']),
    generateCompareColors: jest.fn().mockReturnValue({ left: '#ff0000', right: '#0000ff' }),
    generateTintedGray: jest.fn().mockReturnValue('#cccccc')
}));

// Manual mocks for GAS objects
const mockSetSolidFill = jest.fn();
const mockSetWeight = jest.fn();
const mockSetTransparent = jest.fn();

const mockGetLineFill = jest.fn().mockReturnValue({
    setSolidFill: mockSetSolidFill
});

const mockBorder = {
    getLineFill: mockGetLineFill,
    setWeight: mockSetWeight,
    setTransparent: mockSetTransparent
};

const mockShape = {
    getFill: jest.fn().mockReturnValue({
        setSolidFill: jest.fn(),
        setTransparent: jest.fn()
    }),
    getBorder: jest.fn().mockReturnValue(mockBorder),
    getText: jest.fn().mockReturnValue({
        setText: jest.fn(),
        asString: jest.fn().mockReturnValue(''),
        getRange: jest.fn().mockReturnValue({
            getTextStyle: jest.fn().mockReturnValue({
                setFontSize: jest.fn()
            })
        }),
        getParagraphStyle: jest.fn().mockReturnValue({
            setParagraphAlignment: jest.fn()
        }),
        getTextStyle: jest.fn().mockReturnValue({
            setFontSize: jest.fn().mockReturnThis(),
            setForegroundColor: jest.fn().mockReturnThis(),
            setBold: jest.fn().mockReturnThis(),
            setFontFamily: jest.fn().mockReturnThis()
        })
    }),
    setContentAlignment: jest.fn(),
    setRotation: jest.fn(),
    setEndArrow: jest.fn(),
    setLeft: jest.fn(),
    setTop: jest.fn(),
    setWidth: jest.fn(),
    setHeight: jest.fn(),
    remove: jest.fn() // For placeholder removal
};

const mockTable = {
    getCell: jest.fn().mockReturnValue({
        getText: jest.fn().mockReturnValue({
            setText: jest.fn(),
            getTextStyle: jest.fn().mockReturnValue({
                setBold: jest.fn().mockReturnThis(),
                setFontSize: jest.fn().mockReturnThis(),
                setForegroundColor: jest.fn().mockReturnThis()
            })
        }),
        getFill: jest.fn().mockReturnValue({ setSolidFill: jest.fn() }),
        setContentAlignment: jest.fn()
    }),
    setLeft: jest.fn(),
    setTop: jest.fn(),
    setWidth: jest.fn()
};

const mockLine = {
    getLineFill: mockGetLineFill,
    setWeight: mockSetWeight,
    setEndArrow: jest.fn()
};

const mockPlaceholder = {
    asShape: jest.fn().mockReturnValue(mockShape),
    getLeft: jest.fn().mockReturnValue(0),
    getTop: jest.fn().mockReturnValue(0),
    getWidth: jest.fn().mockReturnValue(960),
    getHeight: jest.fn().mockReturnValue(540),
    remove: jest.fn()
};

const mockSlide = {
    insertShape: jest.fn().mockReturnValue(mockShape),
    insertLine: jest.fn().mockReturnValue(mockLine),
    insertTable: jest.fn().mockReturnValue(mockTable),
    getPlaceholder: jest.fn().mockReturnValue(mockPlaceholder),
    getBackground: jest.fn().mockReturnValue({ setSolidFill: jest.fn() }),
    getPlaceholders: jest.fn().mockReturnValue([]),
    getPageElements: jest.fn().mockReturnValue([]),
    group: jest.fn().mockReturnValue({ getObjectId: jest.fn().mockReturnValue('group-1') })
};

const mockLayout = {
    getRect: jest.fn().mockReturnValue({ left: 0, top: 0, width: 960, height: 540 }),
    pxToPt: jest.fn((px) => px * 0.75) // Simple mock conversion
};

describe('GasDiagramSlideGenerator (Rich Features)', () => {
    beforeEach(() => {
        jest.clearAllMocks();
        (global as any).SlidesApp = {
            ShapeType: {
                RECTANGLE: 'RECTANGLE',
                ELLIPSE: 'ELLIPSE',
                TEXT_BOX: 'TEXT_BOX',
                CHEVRON: 'CHEVRON',
                ROUND_RECTANGLE: 'ROUND_RECTANGLE',
                DONUT: 'DONUT',
                TRAPEZOID: 'TRAPEZOID',
                DOWN_ARROW: 'DOWN_ARROW',
                BENT_ARROW: 'BENT_ARROW'
            },
            LineCategory: { STRAIGHT: 'STRAIGHT' },
            ArrowStyle: { FILL_ARROW: 'FILL_ARROW' },
            PlaceholderType: { TITLE: 'TITLE', CENTERED_TITLE: 'CENTERED_TITLE', BODY: 'BODY' },
            ParagraphAlignment: { CENTER: 'CENTER', START: 'START', LEFT: 'LEFT' },
            ContentAlignment: { MIDDLE: 'MIDDLE' }
        };
        (global as any).Logger = { log: jest.fn() };
    });

    it('should handle "cards" type and draw shapes', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'cards',
            items: [{ title: 'Card 1', desc: 'Desc 1' }, { title: 'Card 2', desc: 'Desc 2' }]
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);

        // Should draw rectangles for cards
        expect(mockSlide.insertShape).toHaveBeenCalled();
        // Should rely on createGrid or similar logic
    });

    it('should handle "kpi" type', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'kpi',
            items: [{ label: 'Sales', value: '100', change: '10%' }]
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);

        expect(mockSlide.insertShape).toHaveBeenCalled();
        // Should verify text setting specifically for KPI values
    });

    it('should handle "table" type using insertTable', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'table',
            headers: ['Col1', 'Col2'],
            rows: [['A', 'B'], ['C', 'D']]
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);

        expect(mockSlide.insertTable).toHaveBeenCalledWith(
            expect.any(Number), // rows
            expect.any(Number)  // cols
        );
    });

    it('should handle "faq" type', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'faq',
            items: [{ q: 'Question?', a: 'Answer.' }]
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);

        // Should draw Q and A boxes
        expect(mockSlide.insertShape).toHaveBeenCalled();
    });

    it('should handle "progress" type', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'progress',
            items: [{ label: 'Task A', percent: 80 }]
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);
        expect(mockSlide.insertShape).toHaveBeenCalled();
    });

    it('should handle "quote" type', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'quote',
            text: 'Vision statement',
            author: 'CEO'
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);
        expect(mockSlide.insertShape).toHaveBeenCalled();
    });

    it('should handle "imageText" type', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'imageText',
            image: 'https://example.com/img.png',
            points: ['Point 1']
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);
        // Should try to insert text and potentially image
        // (Image insertion might be mocked via SlideUtils in real impl, or direct. Verify mock usage)
        const { insertImageFromUrlOrFileId } = require('../../../common/utils/SlideUtils');
        // Actually, existing code might use Utils for image. If proper impl uses the util, we expect it called.
        // We will implement `drawImageText` to use `insertImageFromUrlOrFileId` if possible or similar logic.
    });
});
