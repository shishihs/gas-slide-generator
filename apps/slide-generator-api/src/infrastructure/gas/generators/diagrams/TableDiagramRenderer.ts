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

        // Insert table and position it
        const table = slide.insertTable(numRows, numCols);
        table.setLeft(area.left);
        table.setTop(area.top);
        table.setWidth(area.width);

        const theme = layout.getTheme();

        let rowIndex = 0;
        // Header row
        if (headers.length) {
            for (let c = 0; c < numCols; c++) {
                const cell = table.getCell(0, c);
                cell.getFill().setSolidFill(settings.primaryColor);
                cell.getText().setText(headers[c] || '');
                const style = cell.getText().getTextStyle();
                style.setFontFamily(theme.fonts.family);
                style.setBold(true);
                style.setFontSize(13);
                style.setForegroundColor('#FFFFFF');
                try { cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }
            rowIndex++;
        }

        // Data rows with alternating colors
        rows.forEach((row: any[], rIdx: number) => {
            const isAlt = rIdx % 2 !== 0;
            const rowColor = isAlt ? theme.colors.faintGray : '#FFFFFF';

            for (let c = 0; c < numCols; c++) {
                const cell = table.getCell(rowIndex, c);
                cell.getFill().setSolidFill(rowColor);
                cell.getText().setText(String(row[c] ?? ''));
                const rowStyle = cell.getText().getTextStyle();
                rowStyle.setFontFamily(theme.fonts.family);
                rowStyle.setFontSize(11);
                rowStyle.setForegroundColor(theme.colors.textPrimary);
                try { cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }
            rowIndex++;
        });
    }
}
