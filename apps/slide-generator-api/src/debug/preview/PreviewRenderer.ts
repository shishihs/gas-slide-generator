
import * as fs from 'fs';
import * as path from 'path';
import { VirtualSlide, VirtualShape, VirtualLine } from './VirtualSlide';

export class PreviewRenderer {
    static render(slide: VirtualSlide, outputPath: string) {
        const width = slide.width;
        const height = slide.height;

        let scgContent = '';

        // Sort elements by some z-index if needed, but array order is usually enough
        slide.elements.forEach(el => {
            if (el instanceof VirtualShape) {
                // Determine Shape
                let shapeSvg = '';
                const style = `fill:${el.fillColor === 'TRANSPARENT' ? 'none' : el.fillColor}; stroke:${el.borderColor === 'TRANSPARENT' ? 'none' : el.borderColor}; stroke-width:${el.borderWeight}`;

                if (el.type.includes('RECTANGLE') || el.type.includes('TEXT_BOX')) {
                    shapeSvg = `<rect x="${el.left}" y="${el.top}" width="${el.width}" height="${el.height}" style="${style}" />`;
                } else if (el.type.includes('ELLIPSE')) {
                    const cx = el.left + el.width / 2;
                    const cy = el.top + el.height / 2;
                    const rx = el.width / 2;
                    const ry = el.height / 2;
                    shapeSvg = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" style="${style}" />`;
                } else {
                    // Fallback
                    shapeSvg = `<rect x="${el.left}" y="${el.top}" width="${el.width}" height="${el.height}" style="${style} stroke-dasharray: 4" />`;
                }

                scgContent += shapeSvg;

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

                    const textStyle = `font-family: Arial; font-size: ${el.textStyle.size}pt; fill: ${el.textStyle.color}; font-weight: ${el.textStyle.bold ? 'bold' : 'normal'}`;

                    // Very basic text wrapping or truncation handling could go here, for now just dump it
                    scgContent += `<text x="${tx}" y="${ty}" text-anchor="${textAnchor}" style="${textStyle}">${this.escapeHtml(el.text)}</text>`;
                }

            } else if (el instanceof VirtualLine) {
                const style = `stroke:${el.strokeColor}; stroke-width:${el.weight}`;
                scgContent += `<line x1="${el.x1}" y1="${el.y1}" x2="${el.x2}" y2="${el.y2}" style="${style}" />`;
            }
        });

        const html = `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { background: #f0f0f0; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .slide { background: white; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
    </style>
</head>
<body>
    <div class="slide">
        <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
            ${scgContent}
        </svg>
    </div>
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
