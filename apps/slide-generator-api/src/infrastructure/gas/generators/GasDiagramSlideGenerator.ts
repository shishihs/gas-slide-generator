
import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { CONFIG } from '../../../common/config/SlideConfig';
import {
    setStyledText,
    offsetRect,
    drawBottomBar,
    addCucFooter
} from '../../../common/utils/SlideUtils';

export class GasDiagramSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {
        Logger.log(`Generating Diagram Slide: ${data.layout || data.type}`);

        // Set Title
        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
        if (titlePlaceholder) {
            titlePlaceholder.asShape().getText().setText(data.title || '');
        }

        // Set Subhead
        if (data.subhead) {
            // Basic subhead handling if not already done by standard utils (we'll implement minimal here for diagrams)
            // Or reuse the existing utils. For diagrams, we often need the full body area.
        }

        const type = (data.layout || data.type || '').toLowerCase();

        const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
        const workArea = bodyPlaceholder ?
            { left: bodyPlaceholder.getLeft(), top: bodyPlaceholder.getTop(), width: bodyPlaceholder.getWidth(), height: bodyPlaceholder.getHeight() } :
            layout.getRect('contentSlide.body');

        // switch logic for diagram types
        if (type.includes('timeline')) {
            this.drawTimeline(slide, data, workArea, settings);
        } else if (type.includes('process')) {
            this.drawProcess(slide, data, workArea, settings);
        } else if (type.includes('cycle')) {
            this.drawCycle(slide, data, workArea, settings);
        } else if (type.includes('triangle') || type.includes('pyramid')) {
            this.drawPyramid(slide, data, workArea, settings);
        } else if (type.includes('compare') || type.includes('kaizen')) {
            this.drawComparison(slide, data, workArea, settings);
        } else if (type.includes('stepup') || type.includes('stair')) {
            this.drawStepUp(slide, data, workArea, settings);
        } else {
            // Fallback
            Logger.log('Diagram logic not implemented for type: ' + type);
        }

        if (settings.showBottomBar) {
            drawBottomBar(slide, layout, settings);
        }
        addCucFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }

    private drawTimeline(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any) {
        const milestones = data.milestones || data.items || [];
        if (!milestones.length) return;

        const count = milestones.length;
        const gap = 20;
        const itemWidth = (area.width - (gap * (count - 1))) / count;
        const centerY = area.top + (area.height / 2);

        // Draw Line
        const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, centerY - 2, area.width, 4);
        line.getBorder().setTransparent();
        line.getFill().setSolidFill(CONFIG.COLORS.primary_color);

        milestones.forEach((m: any, i: number) => {
            const x = area.left + (i * (itemWidth + gap));
            const bubble = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x + (itemWidth / 2) - 15, centerY - 15, 30, 30);
            bubble.getFill().setSolidFill('#FFFFFF');
            // Border color: getFill() returns LineFill
            // @ts-ignore
            bubble.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.primary_color);
            bubble.getBorder().setWeight(2);

            // Date/Label
            const dateBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, centerY - 50, itemWidth, 30);
            dateBox.getText().setText(m.date || m.label || '');
            dateBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

            // Description
            const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, centerY + 20, itemWidth, 60);
            descBox.getText().setText(m.state || m.description || '');
            descBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
            descBox.getText().getTextStyle().setFontSize(10);
        });
    }

    private drawProcess(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any) {
        const steps = data.steps || data.items || [];
        if (!steps.length) return;

        const count = steps.length;
        const gap = 10;
        const itemWidth = (area.width - (gap * (count - 1))) / count;

        let x = area.left;
        const y = area.top + 50;
        const h = 100;
        // @ts-ignore
        const accentColor = CONFIG.COLORS.process_arrow || CONFIG.COLORS.primary_color;

        steps.forEach((step: any, i: number) => {
            const shape = slide.insertShape(SlidesApp.ShapeType.CHEVRON, x, y, itemWidth, h);
            shape.getFill().setSolidFill(accentColor);
            shape.getText().setText(typeof step === 'string' ? step : (step.label || step.title || ''));
            shape.getText().getTextStyle().setForegroundColor('#FFFFFF').setBold(true);
            x += itemWidth + gap;
        });
    }

    private drawCycle(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any) {
        // Simple 4-item cycle
        const items = data.items || [];
        const centerX = area.left + area.width / 2;
        const centerY = area.top + area.height / 2;
        const radius = Math.min(area.width, area.height) / 3;
        const secondaryColor = '#666666'; // Fallback for missing config

        items.forEach((item: any, i: number) => {
            const angle = (2 * Math.PI / items.length) * i - (Math.PI / 2); // Start top
            const x = centerX + Math.cos(angle) * (radius * 1.5) - 60;
            const y = centerY + Math.sin(angle) * (radius * 1.5) - 40;

            const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, 120, 80);
            shape.getFill().setSolidFill(secondaryColor);
            shape.getText().setText((item.label || item.title || '') + '\n' + (item.subLabel || item.desc || ''));
            shape.getText().getTextStyle().setForegroundColor('#FFFFFF');

            // Connect to center?
        });

        // Center Core
        const core = slide.insertShape(SlidesApp.ShapeType.DONUT, centerX - radius, centerY - radius, radius * 2, radius * 2);
        core.getFill().setSolidFill(CONFIG.COLORS.primary_color);
        if (data.centerText) {
            const coreText = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerX - 40, centerY - 20, 80, 40);
            coreText.getText().setText(data.centerText);
            coreText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
            coreText.getText().getTextStyle().setBold(true).setFontSize(18).setForegroundColor('#FFFFFF');
        }
    }

    private drawPyramid(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any) {
        const levels = data.levels || data.items || [];
        const count = levels.length;
        const stepHeight = Math.min(area.height, 400) / count;
        const maxWidth = Math.min(area.width, 500);
        const centerX = area.left + area.width / 2;
        const topY = area.top + 20;

        levels.forEach((lvl: any, i: number) => {
            // Draw trapezoids from top to bottom
            // Top level is triangle, others are trapezoids.
            // Calculating width based on level index (0 is top)
            const currentWidth = maxWidth * ((i + 1) / count);
            // const prevWidth = maxWidth * (i / count); // This variable was unused

            const y = topY + (i * stepHeight);
            const x = centerX - (currentWidth / 2);

            // Simplification: Just stack rectangles of varying width for robust rendering for now
            // const shape = slide.insertShape(SlidesApp.ShapeType.TRAPEZOID, x, y, currentWidth, stepHeight); // This line was commented out
            // shape.setRotation(0); // Trapezoid orientation might need adjustment

            // Alternatively, simply use Rectangles for a step-pyramid look as it's cleaner in GAS
            const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, currentWidth, stepHeight - 5);
            rect.getFill().setSolidFill(CONFIG.COLORS.primary_color, 1.0 - (i * 0.15)); // Graduated color with Alpha

            rect.getText().setText((lvl.title || lvl.label || '') + '\n' + (lvl.description || ''));
            rect.getText().getTextStyle().setForegroundColor('#FFFFFF');
        });
    }

    private drawComparison(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any) {
        const leftTitle = data.leftTitle || 'Plan A';
        const rightTitle = data.rightTitle || 'Plan B';

        const gap = 20;
        const colWidth = (area.width - gap) / 2;

        // Left Box
        const leftBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, colWidth, area.height - 50);
        leftBox.getFill().setSolidFill('#F0F0F0');
        leftBox.getText().setText(leftTitle + '\n\n' + (data.leftItems || []).join('\n'));

        // Right Box
        const rightBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left + colWidth + gap, area.top, colWidth, area.height - 50);
        rightBox.getFill().setSolidFill('#E0E0E0');
        rightBox.getText().setText(rightTitle + '\n\n' + (data.rightItems || []).join('\n'));
    }

    private drawStepUp(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any) {
        const items = data.items || [];
        const count = items.length;
        const stepWidth = area.width / count;
        const stepHeight = area.height / count;

        items.forEach((item: any, i: number) => {
            const h = (i + 1) * stepHeight;
            const x = area.left + (i * stepWidth);
            const y = area.top + (area.height - h);

            const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, stepWidth - 5, h);
            shape.getFill().setSolidFill(CONFIG.COLORS.primary_color, 0.5 + (i * 0.1));

            // Text at top of bar
            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, stepWidth - 5, 50);
            textBox.getText().setText(item.title || '');
            textBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
            textBox.getText().getTextStyle().setForegroundColor('#FFFFFF').setBold(true);
        });
    }
}
