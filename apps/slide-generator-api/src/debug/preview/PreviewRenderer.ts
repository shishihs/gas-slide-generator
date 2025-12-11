
import * as fs from 'fs';
import * as path from 'path';
import { VirtualSlide, VirtualShape, VirtualLine, VirtualTable } from './VirtualSlide';

export class PreviewRenderer {
    static render(slides: VirtualSlide[], outputPath: string) {
        let allSvgContent = '';
        const gap = 20;

        slides.forEach((slide, index) => {
            const width = slide.width;
            const height = slide.height;
            let svgContent = '';

            // Sort elements by some z-index if needed
            slide.elements.forEach(el => {
                if (el instanceof VirtualShape) {
                    // Determine Shape
                    let shapeSvg = '';
                    const style = `fill:${el.fillColor === 'TRANSPARENT' ? 'none' : el.fillColor}; stroke:${el.borderColor === 'TRANSPARENT' ? 'none' : el.borderColor}; stroke-width:${el.borderWeight}`;

                    if (el.type.includes('RECTANGLE') || el.type.includes('TEXT_BOX') || el.type.includes('IMAGE')) {
                        // Image treated as rect for now
                        shapeSvg = `<rect x="${el.left}" y="${el.top}" width="${el.width}" height="${el.height}" style="${style}" />`;
                        if (el.type.includes('IMAGE')) {
                            // Add a label for image
                            shapeSvg += `<text x="${el.left + 5}" y="${el.top + 20}" style="font-family:Arial; font-size:10pt; fill:#666">[Image]</text>`;
                        }
                    } else if (el.type.includes('ELLIPSE')) {
                        const cx = el.left + el.width / 2;
                        const cy = el.top + el.height / 2;
                        const rx = el.width / 2;
                        const ry = el.height / 2;
                        shapeSvg = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" style="${style}" />`;
                    } else {
                        // Fallback
                        shapeSvg = `<rect x="${el.left}" y="${el.top}" width="${el.width}" height="${el.height}" style="${style}; stroke-dasharray: 4" />`;
                    }

                    svgContent += shapeSvg;

                    // Text
                    if (el.text) {
                        // Simple text centering/positioning logic
                        let tx = el.left + 5;
                        let ty = el.top + 20;
                        let textAnchor = 'start';

                        if (el.textStyle.align.includes('CENTER')) {
                            tx = el.left + el.width / 2;
                            textAnchor = 'middle';
                        }

                        // Vertical Align helper
                        if (el.contentAlignment === 'MIDDLE') {
                            ty = el.top + el.height / 2 + (el.textStyle.size / 3);
                        }

                        const textStyle = `font-family: Arial; font-size: ${el.textStyle.size}pt; fill: ${el.textStyle.color}; font-weight: ${el.textStyle.bold ? 'bold' : 'normal'}; pointer-events: none;`;

                        svgContent += `<text x="${tx}" y="${ty}" text-anchor="${textAnchor}" style="${textStyle}">${this.escapeHtml(el.text)}</text>`;
                    }

                } else if (el instanceof VirtualLine) {
                    const style = `stroke:${el.strokeColor}; stroke-width:${el.weight}`;
                    svgContent += `<line x1="${el.x1}" y1="${el.y1}" x2="${el.x2}" y2="${el.y2}" style="${style}" />`;
                } else if (el instanceof VirtualTable) {
                    // Render Table using basic rects for cells
                    const rowH = el.height / el.rows; // fallback if dynamic resizing not supported
                    // But VirtualTable doesn't have height/width set by layout manager necessarily, 
                    // it depends on how it was inserted. 
                    // In mock, table just sits at (0,0) with dummy size unless set.
                    // The renderer calls: table.setLeft(), setWidth().
                    // It does NOT set row heights explicitly usually unless auto-calculated.
                    // We'll approximate row height = 30px if not set.

                    const colW = el.width / el.cols;
                    let currentY = el.top;

                    el.cells.forEach((row, rIdx) => {
                        let rowHeight = 30; // Default approximation

                        // Render cells
                        row.forEach((cell, cIdx) => {
                            const x = el.left + (cIdx * colW);
                            const y = currentY;

                            // Cell Background
                            const fillStyle = `fill:${cell.fillColor === 'TRANSPARENT' ? 'none' : cell.fillColor}`;
                            svgContent += `<rect x="${x}" y="${y}" width="${colW}" height="${rowHeight}" style="${fillStyle};" />`;

                            // Cell Text
                            if (cell.text) {
                                const ts = cell.textStyle;
                                const textStyle = `font-family: Arial; font-size: ${ts.size}pt; fill: ${ts.color}; font-weight: ${ts.bold ? 'bold' : 'normal'}; pointer-events: none;`;

                                // Align
                                let tx = x + 5;
                                let ty = y + 20;
                                let anchor = 'start';
                                if (ts.align.includes('CENTER')) { tx = x + colW / 2; anchor = 'middle'; }

                                // Vertical Align (approx)
                                if (cell.contentAlignment === 'MIDDLE') ty = y + rowHeight / 2 + 5;
                                else if (cell.contentAlignment === 'BOTTOM') ty = y + rowHeight - 5;

                                svgContent += `<text x="${tx}" y="${ty}" text-anchor="${anchor}" style="${textStyle}">${this.escapeHtml(cell.text)}</text>`;
                            }

                            // Borders (Simplified: just strokes on the rect? No, borders can be per side)
                            // Top
                            if (cell.borders.top.isVisible) {
                                svgContent += `<line x1="${x}" y1="${y}" x2="${x + colW}" y2="${y}" style="stroke:${cell.borders.top.color}; stroke-width:${cell.borders.top.weight}" />`;
                            }
                            // Bottom
                            if (cell.borders.bottom.isVisible) {
                                svgContent += `<line x1="${x}" y1="${y + rowHeight}" x2="${x + colW}" y2="${y + rowHeight}" style="stroke:${cell.borders.bottom.color}; stroke-width:${cell.borders.bottom.weight}" />`;
                            }
                            // Left/Right ignored for brevity unless vital
                        });
                        currentY += rowHeight;
                    });
                }
            });

            allSvgContent += `
                <div class="slide-container">
                    <div class="slide-label">Slide ${index + 1}</div>
                    <div class="slide">
                        <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
                            ${svgContent}
                        </svg>
                    </div>
                </div>
            `;
        });

        const html = `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { background: #f0f0f0; padding: 20px; font-family: sans-serif; }
        .slide-container { margin-bottom: 40px; display: flex; flex-direction: column; align-items: center; }
        .slide-label { margin-bottom: 10px; font-weight: bold; color: #666; }
        .slide { background: white; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
    </style>
</head>
<body>
    ${allSvgContent}
</body>
</html>
        `;

        fs.writeFileSync(outputPath, html);
        console.log(`Preview generated at: ${outputPath}`);
    }

    private static escapeHtml(text: string) {
        return text
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }
}
