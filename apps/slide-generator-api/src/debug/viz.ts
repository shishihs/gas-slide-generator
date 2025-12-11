import * as fs from 'fs';
import * as path from 'path';
import { VirtualSlide } from './preview/VirtualSlide';
import { PreviewRenderer } from './preview/PreviewRenderer';
import { GasDiagramSlideGenerator } from '../infrastructure/gas/generators/GasDiagramSlideGenerator';
import { GasTitleSlideGenerator } from '../infrastructure/gas/generators/GasTitleSlideGenerator';
import { GasSectionSlideGenerator } from '../infrastructure/gas/generators/GasSectionSlideGenerator';
import { GasContentSlideGenerator } from '../infrastructure/gas/generators/GasContentSlideGenerator';
import { LayoutManager } from '../common/utils/LayoutManager';
import { DEFAULT_THEME, AVAILABLE_THEMES } from '../common/config/DefaultTheme';

// --- MOCK GLOBALS ---
(global as any).SlidesApp = {
    ShapeType: {
        RECTANGLE: 'RECTANGLE',
        ELLIPSE: 'ELLIPSE',
        TEXT_BOX: 'TEXT_BOX',
        ROUND_RECTANGLE: 'ROUND_RECTANGLE',
        LINE: 'LINE'
    },
    LineCategory: { STRAIGHT: 'STRAIGHT' },
    PlaceholderType: {
        BODY: 'BODY',
        TITLE: 'TITLE',
        CENTERED_TITLE: 'CENTERED_TITLE',
        SUBTITLE: 'SUBTITLE',
        OBJECT: 'OBJECT',
        PICTURE: 'PICTURE'
    },
    ParagraphAlignment: { CENTER: 'CENTER', LEFT: 'LEFT', START: 'START', END: 'END', JUSTIFIED: 'JUSTIFIED' },
    ContentAlignment: { MIDDLE: 'MIDDLE', TOP: 'TOP', BOTTOM: 'BOTTOM' }
};

(global as any).Logger = {
    log: (msg: any) => console.log(`[GAS Log] ${msg}`)
};
// --------------------

// Load Data
const jsonPath = path.resolve(__dirname, '../test-data/comprehensive-test.json');
const rawData = fs.readFileSync(jsonPath, 'utf8');
let slideDataList = JSON.parse(rawData);
let themeName = 'Green';

if (!Array.isArray(slideDataList) && slideDataList.slides) {
    themeName = slideDataList.theme || themeName;
    slideDataList = slideDataList.slides;
}

// Generators
const titleGenerator = new GasTitleSlideGenerator(null);
const sectionGenerator = new GasSectionSlideGenerator(null);
const contentGenerator = new GasContentSlideGenerator(null);
const diagramGenerator = new GasDiagramSlideGenerator(null);

const selectedTheme = AVAILABLE_THEMES[themeName] || DEFAULT_THEME;
console.log(`Using Theme: ${themeName}`);

const layout = new LayoutManager(960, 540, selectedTheme);
const settings = {
    primaryColor: selectedTheme.colors.primary,
    secondaryColor: selectedTheme.colors.deepPrimary,
    fonts: { body: 'Arial' }
};

const generatedSlides: VirtualSlide[] = [];

console.log(`Found ${slideDataList.length} slides in mock data.`);

slideDataList.forEach((data: any, index: number) => {
    // Skip comments or invalid entries if any
    if (!data.type) return;

    console.log(`Processing Slide ${index + 1}: ${data.type}`);
    const slide = new VirtualSlide();

    // Routing Logic (Simulating GasSlideRepository)
    const type = (data.type || '').toLowerCase();

    try {
        if (type === 'title') {
            titleGenerator.generate(slide as any, data, layout, index + 1, settings);
        } else if (type === 'section') {
            sectionGenerator.generate(slide as any, data, layout, index + 1, settings);
        } else if (['content', 'agenda', 'unknown'].includes(type)) {
            contentGenerator.generate(slide as any, data, layout, index + 1, settings);
        } else {
            // Assume parsing or diagram type
            diagramGenerator.generate(slide as any, data, layout, index + 1, settings);
        }
        generatedSlides.push(slide);
    } catch (e) {
        console.error(`Failed to generate slide ${index + 1} (${type}):`, e);
    }
});

console.log(`Rendering preview for ${generatedSlides.length} slides...`);
PreviewRenderer.render(generatedSlides, 'preview.html');
