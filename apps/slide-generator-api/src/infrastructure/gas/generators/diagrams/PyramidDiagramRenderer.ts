import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generatePyramidColors } from '../../../../common/utils/ColorUtils';

export class PyramidDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const levels = data.levels || data.items || [];
        if (!levels.length) return;

        // Reverse levels so the first item is at the top (Apex) or bottom (Base)? 
        // Typically Pyramid Principle: Main Message (Top) -> Supporting (Bottom).
        // Let's assume index 0 is Top.

        const count = levels.length;
        const pyramidH = Math.min(area.height, 400);
        const pyramidW = Math.min(area.width * 0.6, 500); // Pyramid takes 60% width

        const centerX = area.left + (area.width / 2);
        const topY = area.top + (area.height - pyramidH) / 2;
        const bottomY = topY + pyramidH;

        // Gap between levels
        const gap = 4;
        const levelHeight = (pyramidH - (gap * (count - 1))) / count;

        levels.forEach((level: any, index: number) => {
            // Calculate Trapezoid Coordinates
            // Top width factor (0 at apex, 1 at base)
            const topRatio = index / count;
            const bottomRatio = (index + 1) / count;

            const segmentTopY = topY + (index * (levelHeight + gap));
            // Width at segment top
            const segmentTopW = pyramidW * topRatio;
            // Width at segment bottom
            const segmentBottomW = pyramidW * bottomRatio;

            // To verify: Index 0 (Top) -> topRatio=0 (Point), bottomRatio=1/3.

            // Draw shape using Builder? No, standard TRAPEZOID shape doesn't allow easy ratio control in GAS.
            // But we can approximate with basic shapes or standard ISOSCELES_TRIANGLE masked?
            // Actually, simply stacking trapezoids is hard in standard Slides API without FreeformBuilder (which GAS doesn't fully expose easily).
            // Workaround: Overlapping triangles? Or just simple Rectangles of varying width (Step Pyramid)?
            // Let's do a "Triangle" layout where we place text boxes, and maybe a large background Triangle behind everything?
            // Or better: Use 'ShapeType.TRAPEZOID' is fixed ratio? No.

            // Layout Strategy: Draw one big Triangle background, then white lines to slice it?
            // Yes, easiest.
        });

        // 1. Draw Main Triangle Background
        // Use theme primary color
        // Note: Google Apps Script uses 'TRIANGLE' typically for isosceles style in standard set
        const mainTriangle = slide.insertShape(SlidesApp.ShapeType.TRIANGLE, centerX - pyramidW / 2, topY, pyramidW, pyramidH);
        mainTriangle.getFill().setSolidFill(settings.primaryColor);
        mainTriangle.getBorder().setTransparent();

        // 2. Draw White Lines to "slice" levels
        for (let i = 1; i < count; i++) {
            const lineY = topY + (i * (pyramidH / count));
            // Calculate width at this Y
            // Linear interpolation: y goes 0..H, width goes 0..W
            const ratio = i / count;
            const widthAtY = pyramidW * ratio;

            // The line actually needs to be white rectangle to "erase" or just a line
            // A line is fine.
            const line = slide.insertLine(
                SlidesApp.LineCategory.STRAIGHT,
                centerX - widthAtY / 2 + 2, // Slight indent
                lineY,
                centerX + widthAtY / 2 - 2,
                lineY
            );
            line.getLineFill().setSolidFill('#FFFFFF');
            line.setWeight(2);
        }

        // 3. Labels (Outside the pyramid, connected by lines)
        levels.forEach((level: any, index: number) => {
            const y = topY + (index * (pyramidH / count)) + (levelHeight / 2) - 10; // Center of segment

            // Label X Position: Right side of pyramid
            const ratioAtCenter = (index + 0.5) / count;
            const pyramidEdgeX = centerX + (pyramidW * ratioAtCenter) / 2;

            const lineStartX = pyramidEdgeX + 10;
            const lineEndX = lineStartX + 30;
            const textX = lineEndX + 10;

            // Connector Line
            const conn = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineStartX, y + 10, lineEndX, y + 10);
            conn.getLineFill().setSolidFill(settings.text_primary || '#333333');
            conn.setWeight(1);

            // Text Box
            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textX, y - 10, 200, 50);

            const title = level.label || level.title || '';
            const sub = level.subLabel || level.desc || '';
            const textContent = `${title}\n${sub}`;

            const textRange = textBox.getText();
            textRange.setText(textContent);

            // Style
            const titleRange = textRange.getRange(0, title.length);
            titleRange.getTextStyle().setBold(true).setFontSize(16).setForegroundColor(settings.text_primary || '#333333');

            if (sub) {
                const subRange = textRange.getRange(title.length + 1, sub.length);
                subRange.getTextStyle().setBold(false).setFontSize(12).setForegroundColor(settings.ghost_gray || '#666666');
            }
        });
    }
}
