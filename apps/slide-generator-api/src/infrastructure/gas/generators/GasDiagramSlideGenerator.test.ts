import { vi, describe, it, expect, beforeEach, Mock } from 'vitest';
import { GasDiagramSlideGenerator } from './GasDiagramSlideGenerator';
import { setupGasGlobals } from '../../../test/gas-mocks';

// Mock Dependencies
vi.mock('../../../common/utils/SlideUtils', () => ({
    setStyledText: vi.fn(),
    offsetRect: vi.fn(),
    addFooter: vi.fn(),
    drawArrowBetweenRects: vi.fn(),
    setBoldTextSize: vi.fn()
}));

vi.mock('../../../common/utils/ColorUtils', () => ({
    generateProcessColors: vi.fn().mockReturnValue(['#ff0000', '#00ff00', '#0000ff']),
    generateTimelineCardColors: vi.fn().mockReturnValue(['#ff0000', '#00ff00', '#0000ff']),
    generatePyramidColors: vi.fn().mockReturnValue(['#ff0000', '#00ff00', '#0000ff']),
    generateCompareColors: vi.fn().mockReturnValue({ left: '#ff0000', right: '#0000ff' }),
    generateTintedGray: vi.fn().mockReturnValue('#cccccc')
}));

// Manual mocks for GAS objects
const mockSetSolidFill = vi.fn();
const mockSetWeight = vi.fn();
const mockSetTransparent = vi.fn();

const mockGetLineFill = vi.fn().mockReturnValue({
    setSolidFill: mockSetSolidFill
});

const mockBorder = {
    getLineFill: mockGetLineFill,
    setWeight: mockSetWeight,
    setTransparent: mockSetTransparent
};

const mockShape = {
    getFill: vi.fn().mockReturnValue({
        setSolidFill: vi.fn(),
        setTransparent: vi.fn()
    }),
    getBorder: vi.fn().mockReturnValue(mockBorder),
    getText: vi.fn().mockReturnValue({
        setText: vi.fn(),
        asString: vi.fn().mockReturnValue(''),
        getRange: vi.fn().mockReturnValue({
            getTextStyle: vi.fn().mockReturnValue({
                setFontSize: vi.fn()
            })
        }),
        getParagraphStyle: vi.fn().mockReturnValue({
            setParagraphAlignment: vi.fn()
        }),
        getTextStyle: vi.fn().mockReturnValue({
            setFontSize: vi.fn().mockReturnThis(),
            setForegroundColor: vi.fn().mockReturnThis(),
            setBold: vi.fn().mockReturnThis(),
            setFontFamily: vi.fn().mockReturnThis()
        })
    }),
    setContentAlignment: vi.fn(),
    setRotation: vi.fn(),
    setEndArrow: vi.fn()
};

// Line Mock
const mockLine = {
    getLineFill: mockGetLineFill,
    setWeight: mockSetWeight,
    setEndArrow: vi.fn()
};

const mockPlaceholder = {
    asShape: vi.fn().mockReturnValue(mockShape),
    getLeft: vi.fn().mockReturnValue(0),
    getTop: vi.fn().mockReturnValue(0),
    getWidth: vi.fn().mockReturnValue(960),
    getHeight: vi.fn().mockReturnValue(540)
};

const mockSlide = {
    insertShape: vi.fn().mockReturnValue(mockShape),
    insertLine: vi.fn().mockReturnValue(mockLine),
    getPlaceholder: vi.fn().mockReturnValue(mockPlaceholder),
    getPlaceholders: vi.fn().mockReturnValue([]),
    getBackground: vi.fn().mockReturnValue({ setSolidFill: vi.fn() }),
    getPageElements: vi.fn().mockReturnValue([]),
    group: vi.fn().mockReturnValue({ getObjectId: vi.fn().mockReturnValue('group-1') })
};

const mockLayout = {
    getRect: vi.fn().mockReturnValue({ left: 0, top: 0, width: 960, height: 540 }),
    pxToPt: vi.fn((px) => px * 0.75), // Simple mock conversion
    getTheme: vi.fn().mockReturnValue({
        colors: {
            primary: '#4285F4',
            textPrimary: '#000000',
            neutralGray: '#888888',
            backgroundGray: '#F1F3F4',
            laneBorder: '#E0E0E0',
            cardBorder: '#CCCCCC',
            ghostGray: '#F5F5F5',
            faintGray: '#F8F9FA',
            textSmallFont: '#555555'
        },
        fonts: {
            family: 'Roboto',
            sizes: {
                title: 32,
                body: 14,
                laneTitle: 18
            }
        },
        diagram: {
            laneGapPx: 10,
            lanePadPx: 10,
            laneTitleHeightPx: 40,
            cardGapPx: 10,
            cardMinHeightPx: 60,
            cardMaxHeightPx: 120,
            arrowHeightPx: 20,
            arrowGapPx: 5
        }
    })
};

describe('GasDiagramSlideGenerator', () => {
    beforeEach(() => {
        vi.clearAllMocks();
        setupGasGlobals();
    });

    it('drawTimeline should generate correct requests', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'timeline',
            milestones: [
                { date: '2025', label: 'Test Milestone' }
            ]
        };
        const settings = { primaryColor: '#4285F4', showBottomBar: true };

        // Execute
        const results = generator.generate('slide-id', data, mockLayout as any, 1, settings);

        // Verification
        expect(Array.isArray(results)).toBe(true);
        // Should have created lines (axis) and shapes (cards)
        const shapeReq = results.find(r => r.createShape && r.createShape.shapeType === 'ROUND_RECTANGLE');
        const lineReq = results.find(r => r.createLine);

        expect(shapeReq).toBeDefined();
        expect(lineReq).toBeDefined();
    });

    it('drawProcess should generate process requests', () => {
        const generator = new GasDiagramSlideGenerator(null);
        const data = {
            type: 'process',
            steps: ['Step 1', 'Step 2']
        };
        const settings = { primaryColor: '#4285F4' };

        const results = generator.generate('slide-id', data, mockLayout as any, 1, settings);

        expect(Array.isArray(results)).toBe(true);
        const shapeReq = results.find(r => r.createShape && r.createShape.shapeType === 'CHEVRON');
        expect(shapeReq).toBeDefined();
    });
});
