
import { VirtualSlide } from './preview/VirtualSlide';
import { PreviewRenderer } from './preview/PreviewRenderer';
import { GasDiagramSlideGenerator } from '../infrastructure/gas/generators/GasDiagramSlideGenerator';
import { LayoutManager } from '../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../common/config/DefaultTheme';

// --- MOCK GLOBALS ---
(global as any).SlidesApp = {
    ShapeType: {
        RECTANGLE: 'RECTANGLE',
        ELLIPSE: 'ELLIPSE',
        TEXT_BOX: 'TEXT_BOX',
        ROUND_RECTANGLE: 'ROUND_RECTANGLE'
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
    ParagraphAlignment: { CENTER: 'CENTER', LEFT: 'LEFT', START: 'START' },
    ContentAlignment: { MIDDLE: 'MIDDLE' }
};

(global as any).Logger = {
    log: (msg: any) => console.log(`[GAS Log] ${msg}`)
};
// --------------------

// Setup
const slide = new VirtualSlide();
// GasDiagramSlideGenerator expects a slide object that implements the GAS interface. 
// VirtualSlide implements the subset identifying methods used.

const generator = new GasDiagramSlideGenerator(null);
const layout = new LayoutManager(960, 540, DEFAULT_THEME);

// Input Data
const data = {
    type: 'timeline',
    title: 'Project Timeline 2025',
    milestones: [
        { date: 'Q1', label: 'Research Phase' },
        { date: 'Q2', label: 'Prototyping' },
        { date: 'Q3', label: 'Development' },
        { date: 'Q4', label: 'Launch' }
    ]
};

const settings = {
    primaryColor: '#4285F4',
    secondaryColor: '#EA4335',
    fonts: { body: 'Arial' }
};

console.log('Generating slide...');
generator.generate(slide as any, data, layout, 1, settings);

console.log('Rendering preview...');
PreviewRenderer.render(slide, 'preview.html');
