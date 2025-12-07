export const MockSlidesApp = {
    create: jest.fn(),
    openById: jest.fn(),
    PredefinedLayout: {
        BLANK: 'BLANK',
        TITLE_AND_BODY: 'TITLE_AND_BODY',
    },
    ShapeType: {
        TEXT_BOX: 'TEXT_BOX',
        RECTANGLE: 'RECTANGLE',
    },
    ParagraphAlignment: {
        CENTER: 'CENTER',
    },
    ContentAlignment: {
        MIDDLE: 'MIDDLE',
    },
};

export const MockDriveApp = {
    getFileById: jest.fn(),
};

export const MockPropertiesService = {
    getScriptProperties: jest.fn().mockReturnValue({
        getProperty: jest.fn(),
    }),
};

// Global polyfills if needed (simulating GAS environment)
(global as any).SlidesApp = MockSlidesApp;
(global as any).DriveApp = MockDriveApp;
(global as any).PropertiesService = MockPropertiesService;
