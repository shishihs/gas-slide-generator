
import { GasDiagramSlideGenerator } from './GasDiagramSlideGenerator';

// Mock Dependencies
jest.mock('../../../common/utils/SlideUtils', () => ({
    setStyledText: jest.fn(),
    offsetRect: jest.fn(),
    addFooter: jest.fn(),
    drawArrowBetweenRects: jest.fn(),
    setBoldTextSize: jest.fn()
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
    setEndArrow: jest.fn()
};

// Line Mock (different from Shape in some ways, but sharing similar methods for simplicity here)
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
    getHeight: jest.fn().mockReturnValue(540)
};

const mockSlide = {
    insertShape: jest.fn().mockReturnValue(mockShape),
    insertLine: jest.fn().mockReturnValue(mockLine),
    getPlaceholder: jest.fn().mockReturnValue(mockPlaceholder),
    getPlaceholders: jest.fn().mockReturnValue([]),
    getBackground: jest.fn().mockReturnValue({ setSolidFill: jest.fn() }),
    getPageElements: jest.fn().mockReturnValue([]),
    group: jest.fn().mockReturnValue({ getObjectId: jest.fn().mockReturnValue('group-1') })
};

const mockLayout = {
    getRect: jest.fn().mockReturnValue({ left: 0, top: 0, width: 960, height: 540 }),
    pxToPt: jest.fn((px) => px * 0.75) // Simple mock conversion
};

describe('GasDiagramSlideGenerator', () => {
    beforeEach(() => {
        jest.clearAllMocks();

        // Mock Global SlidesApp
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
            LineCategory: {
                STRAIGHT: 'STRAIGHT'
            },
            ArrowStyle: {
                FILL_ARROW: 'FILL_ARROW'
            },
            PlaceholderType: {
                TITLE: 'TITLE',
                CENTERED_TITLE: 'CENTERED_TITLE',
                BODY: 'BODY'
            },
            ParagraphAlignment: {
                CENTER: 'CENTER',
                START: 'START', // Fixed from LEFT
                LEFT: 'LEFT' // Kept for backward compat if code still uses it anywhere else? No, strictly mocking.
            },
            ContentAlignment: {
                MIDDLE: 'MIDDLE'
            }
        };

        // Mock Logger
        (global as any).Logger = {
            log: jest.fn()
        };
    });

    it('drawTimeline should correctly set border color using getLineFill() and use ColorUtils', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'timeline',
            milestones: [
                { date: '2025', label: 'Test Milestone' }
            ]
        };
        const settings = { primaryColor: '#4285F4', showBottomBar: true };

        // Execute
        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);

        // Verification
        expect(mockSlide.insertShape).toHaveBeenCalled();

        // Ensure color generation was called
        const { generateTimelineCardColors } = require('../../../common/utils/ColorUtils');
        expect(generateTimelineCardColors).toHaveBeenCalledWith('#4285F4', 1);

        // Check if getLineFill was accessed (border color setting)
        // In the new implementation we set header shape border
        expect(mockGetLineFill).toHaveBeenCalled();

        // Also verify usage of SlideUtils
        const { setStyledText } = require('../../../common/utils/SlideUtils');
        expect(setStyledText).toHaveBeenCalled();
    });

    it('drawProcess should generate process steps', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'process',
            steps: ['Step 1', 'Step 2']
        };
        const settings = { primaryColor: '#4285F4' };

        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);

        const { generateProcessColors } = require('../../../common/utils/ColorUtils');
        expect(generateProcessColors).toHaveBeenCalledWith('#4285F4', 2);
        expect(mockSlide.insertShape).toHaveBeenCalled();
    });
});
