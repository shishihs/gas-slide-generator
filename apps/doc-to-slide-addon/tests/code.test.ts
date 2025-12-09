
import * as fs from 'fs';
import * as path from 'path';
import * as vm from 'vm';
import * as api from 'typescript';

// Mock GAS Globals
const mockPropertiesService = {
    getScriptProperties: jest.fn().mockReturnValue({
        getProperty: jest.fn().mockImplementation((key) => {
            if (key === 'TEMPLATE_SLIDE_ID') return 'https://docs.google.com/presentation/d/mock-template-id';
            return null; // For others
        })
    })
};

const mockDriveApp = {
    getFileById: jest.fn().mockReturnValue({
        getName: jest.fn().mockReturnValue('Template File'),
        makeCopy: jest.fn().mockReturnValue({
            getId: jest.fn().mockReturnValue('new-copy-id')
        })
    })
};

const mockSlideGeneratorApi = {
    generateSlides: jest.fn().mockReturnValue({
        success: true,
        url: 'https://docs.google.com/presentation/d/new-copy-id'
    })
};

const mockLogger = {
    log: jest.fn()
};

// Define global context
const context = {
    PropertiesService: mockPropertiesService,
    DriveApp: mockDriveApp,
    SlideGeneratorApi: mockSlideGeneratorApi,
    Logger: mockLogger,
    console: console
};

describe('Code.ts Logic Verification', () => {
    let scriptContext: any;

    beforeAll(() => {
        const codePath = path.join(__dirname, '../src/Code.ts');
        const tsCodeContent = fs.readFileSync(codePath, 'utf8');

        // Transpile TS to JS
        const jsContent = api.transpileModule(tsCodeContent, {
            compilerOptions: { module: api.ModuleKind.CommonJS }
        }).outputText;

        // We need to execute the Code.ts content in a VM to populate functions
        vm.createContext(context);
        vm.runInContext(jsContent, context);
        scriptContext = context;
    });

    it('generateSlidesFromJson should orchestration creation correctly', () => {
        const jsonData = {
            title: "Test Pres",
            slides: [{ title: "S1" }]
        };

        // Call the function from the context
        const result = scriptContext.generateSlidesFromJson(jsonData);

        expect(result).toEqual({
            success: true,
            slideUrl: 'https://docs.google.com/presentation/d/new-copy-id'
        });

        // Verify Interactions
        // 1. Get Template ID
        expect(mockPropertiesService.getScriptProperties().getProperty).toHaveBeenCalledWith('TEMPLATE_SLIDE_ID');

        // 2. Get File and Copy
        expect(mockDriveApp.getFileById).toHaveBeenCalledWith('mock-template-id');
        expect(mockDriveApp.getFileById().makeCopy).toHaveBeenCalledWith('Test Pres');

        // 3. Call Library
        expect(mockSlideGeneratorApi.generateSlides).toHaveBeenCalledWith({
            title: "Test Pres",
            slides: jsonData.slides,
            theme: "SIMPLE_LIGHT",
            templateId: "mock-template-id",
            destinationId: "new-copy-id"
        });
    });

    it('should handle library missing error', () => {
        // Modify the EXISTING context to remove the library
        const originalApi = scriptContext.SlideGeneratorApi;
        scriptContext.SlideGeneratorApi = undefined;

        expect(() => {
            scriptContext.generateSlidesFromJson({ slides: [] });
        }).toThrow('SlideGeneratorApi library not found');

        // Restore for other tests if any
        scriptContext.SlideGeneratorApi = originalApi;
    });
});
