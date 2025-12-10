
import { GasDiagramSlideGenerator } from './GasDiagramSlideGenerator';

// Manual mocks for dependent calls
const mockSetSolidFill = jest.fn();
const mockSetWeight = jest.fn();
const mockSetTransparent = jest.fn();

const mockGetLineFill = jest.fn().mockReturnValue({
    setSolidFill: mockSetSolidFill
});

const mockBorder = {
    getLineFill: mockGetLineFill,
    setWeight: mockSetWeight,
    setTransparent: mockSetTransparent,
    // For coverage of previous incorrect implementation attempts, we ensure these don't exist or we can mock them if we want to ensure they ARE NOT called.
    // However, JS allows existing properties.
};

// Shape Mock
const mockShape = {
    getFill: jest.fn().mockReturnValue({
        setSolidFill: jest.fn()
    }),
    getBorder: jest.fn().mockReturnValue(mockBorder),
    insertShape: jest.fn(),
    getText: jest.fn().mockReturnValue({
        setText: jest.fn(),
        getParagraphStyle: jest.fn().mockReturnValue({
            setParagraphAlignment: jest.fn()
        }),
        getTextStyle: jest.fn().mockReturnValue({
            setFontSize: jest.fn().mockReturnThis(),
            setForegroundColor: jest.fn().mockReturnThis(),
            setBold: jest.fn().mockReturnThis(),
            setFontFamily: jest.fn().mockReturnThis()
        })
    })
};

const mockSlide = {
    insertShape: jest.fn().mockReturnValue(mockShape),
    getPlaceholder: jest.fn().mockReturnValue(null),
    getBackground: jest.fn().mockReturnValue({ setSolidFill: jest.fn() })
};

const mockLayout = {
    getRect: jest.fn().mockReturnValue({ left: 0, top: 0, width: 960, height: 540 })
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
            },
            PlaceholderType: {
                TITLE: 'TITLE',
                CENTERED_TITLE: 'CENTERED_TITLE',
                BODY: 'BODY'
            },
            ParagraphAlignment: {
                CENTER: 'CENTER'
            }
        };

        // Mock Logger
        (global as any).Logger = {
            log: jest.fn()
        };
    });

    it('drawTimeline should correctly set border color using getLineFill()', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'timeline',
            milestones: [
                { date: '2025', description: 'Test Milestone' }
            ]
        };
        const settings = {};

        // Execute
        generator.generate(mockSlide as any, data, mockLayout as any, 1, settings);

        // Verification
        expect(mockSlide.insertShape).toHaveBeenCalled();

        // Check if getLineFill was accessed
        expect(mockGetLineFill).toHaveBeenCalled();

        // Check if setSolidFill was called on the LineFill object
        // We know CONFIG.COLORS.primary_color is '#4285F4'
        expect(mockSetSolidFill).toHaveBeenCalledWith('#4285F4');
    });
});
