import { vi, Mock } from 'vitest';

/**
 * Creates a type-safe mock for Google Apps Script SlidesApp.
 * This includes enums and basic mock functions.
 */
export const createMockSlidesApp = () => ({
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
        SUBTITLE: 'SUBTITLE',
        BODY: 'BODY',
        OBJECT: 'OBJECT',
        PICTURE: 'PICTURE'
    },
    ParagraphAlignment: {
        START: 'START',
        CENTER: 'CENTER',
        END: 'END',
        JUSTIFIED: 'JUSTIFIED',
        LEFT: 'LEFT' // For backward compatibility if needed, though START is preferred
    },
    ContentAlignment: {
        TOP: 'TOP',
        MIDDLE: 'MIDDLE',
        BOTTOM: 'BOTTOM'
    },
    create: vi.fn(),
    openById: vi.fn(),
    getActivePresentation: vi.fn()
});

/**
 * Creates a mock Logger.
 */
export const createMockLogger = () => ({
    log: vi.fn(),
    clear: vi.fn(),
    getLog: vi.fn()
});

/**
 * Creates a mock DriveApp.
 */
export const createMockDriveApp = () => ({
    getFileById: vi.fn(),
    createFile: vi.fn(),
    getFolderById: vi.fn()
});

/**
 * Creates a mock Slides (Advanced Service).
 */
export const createMockSlides = () => ({
    Presentations: {
        batchUpdate: vi.fn()
    }
});

/**
 * Helper to setup global GAS environment mocks in tests.
 */
export const setupGasGlobals = () => {
    const slidesApp = createMockSlidesApp();
    const logger = createMockLogger();
    const driveApp = createMockDriveApp();
    const slides = createMockSlides();

    // Use globalThis to safely access global scope in Node environment
    (globalThis as any).SlidesApp = slidesApp;
    (globalThis as any).Logger = logger;
    (globalThis as any).DriveApp = driveApp;
    (globalThis as any).Slides = slides;

    return { slidesApp, logger, driveApp, slides };
};

export type MockSlidesApp = ReturnType<typeof createMockSlidesApp>;
