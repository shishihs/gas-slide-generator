import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class TableDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const headers = data.headers || [];
        const rows = data.rows || [];

        // Determine column count: prefer rows[0].length as it might have more columns than headers
        const numCols = Math.max(
            headers.length,
            rows.length > 0 && Array.isArray(rows[0]) ? rows[0].length : 0
        );
        const numRows = rows.length + (headers.length ? 1 : 0);

        Logger.log(`Table: ${numRows} rows x ${numCols} cols, headers: ${JSON.stringify(headers)}`);
        Logger.log(`Table area: left=${area.left}, top=${area.top}, width=${area.width}, height=${area.height}`);

        if (numRows === 0 || numCols === 0) return;

        // Create Table with minimal structure
        const table = slide.insertTable(numRows, numCols);
        table.setLeft(area.left);
        table.setTop(area.top);
        table.setWidth(area.width);

        const theme = layout.getTheme();

        // Iterate through all cells to set base refined style
        // Clean typography, no heavy borders initially
        for (let r = 0; r < numRows; r++) {
            for (let c = 0; c < numCols; c++) {
                const cell: any = table.getCell(r, c);
                cell.getFill().setTransparent(); // Airy look

                // Borders: Only horizontal lines
                cell.getBorderRight().setTransparent();
                cell.getBorderLeft().setTransparent();
                cell.getBorderTop().setTransparent();

                // Bottom border: subtle
                const borderBottom = cell.getBorderBottom();
                borderBottom.setSolidFill(theme.colors.ghostGray);
                borderBottom.setWeight(1);
            }
        }

        let rowIndex = 0;
        // Header row
        if (headers.length) {
            for (let c = 0; c < numCols; c++) {
                const cell: any = table.getCell(0, c);
                // Editorial Header: No solid background fill, just strong text and a thicker bottom line
                cell.getText().setText(headers[c] || '');
                const style = cell.getText().getTextStyle();
                style.setFontFamily(theme.fonts.family);
                style.setBold(true);
                style.setFontSize(14); // Slightly larger
                style.setForegroundColor(settings.primaryColor); // Colored text instead of white on color

                // Stronger separator for header
                cell.getBorderBottom().setSolidFill(settings.primaryColor);
                cell.getBorderBottom().setWeight(3); // Thicker for clear hierarchy

                try { cell.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch (e) { } // Align bottom for header
            }
            rowIndex++;
        }

        // Data rows
        rows.forEach((row: any[], rIdx: number) => {
            for (let c = 0; c < numCols; c++) {
                const cell: any = table.getCell(rowIndex, c);
                cell.getText().setText(String(row[c] ?? ''));
                const rowStyle = cell.getText().getTextStyle();
                rowStyle.setFontFamily(theme.fonts.family);
                rowStyle.setFontSize(12);
                rowStyle.setForegroundColor(theme.colors.textPrimary);
                try { cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }
            rowIndex++;
        });
    }
}
