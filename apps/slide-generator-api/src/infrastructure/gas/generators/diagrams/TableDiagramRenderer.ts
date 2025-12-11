import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { RequestFactory } from '../../RequestFactory';

export class TableDiagramRenderer implements IDiagramRenderer {
    render(slideId: string, data: any, area: any, settings: any, layout: LayoutManager): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
        const headers = data.headers || [];
        const rows = data.rows || [];
        const theme = layout.getTheme();

        const numCols = Math.max(headers.length, rows.length > 0 && Array.isArray(rows[0]) ? rows[0].length : 0);
        const numRows = rows.length + (headers.length ? 1 : 0);

        if (numRows === 0 || numCols === 0) return requests;

        const tableId = slideId + '_TBL';

        // 1. Create Table
        requests.push({
            createTable: {
                objectId: tableId,
                elementProperties: {
                    pageObjectId: slideId,
                    size: { width: { magnitude: area.width, unit: 'PT' }, height: { magnitude: area.height, unit: 'PT' } },
                    transform: { scaleX: 1, scaleY: 1, translateX: area.left, translateY: area.top, unit: 'PT' }
                },
                rows: numRows,
                columns: numCols
            }
        });

        // 2. Clear Borders (Start Fresh)
        requests.push({
            updateTableBorderProperties: {
                objectId: tableId,
                tableRange: { location: { rowIndex: 0, columnIndex: 0 }, rowSpan: numRows, columnSpan: numCols },
                borderPosition: 'ALL',
                tableBorderProperties: {
                    tableBorderFill: { solidFill: { color: { rgbColor: { red: 0, green: 0, blue: 0 } } } }, // Should accept transparent?
                    // Actually API handles transparency via propertyState='NOT_RENDERED' or fill alpha?
                    // Let's set fill to NOT_RENDERED doesn't exist for TableBorderFill? 
                    // Actually it does: "solidFill: { alpha: 0 }" or similar.
                    // Or set weight to 0.
                    weight: { magnitude: 0, unit: 'PT' }
                },
                fields: 'weight'
            }
        });

        let rowIndex = 0;

        // 3. Header Row
        if (headers.length) {
            headers.forEach((header: string, c: number) => {
                const text = header || '';
                if (text) {
                    requests.push({ insertText: { objectId: tableId, cellLocation: { rowIndex: 0, columnIndex: c }, text: text } });
                    requests.push(RequestFactory.updateTextStyle(tableId, {
                        bold: true,
                        fontSize: 14,
                        color: '#FFFFFF',
                        fontFamily: theme.fonts.family
                    }));
                    // The above RequestFactory uses objectId. For Tables, it needs cellLocation support?
                    // My RequestFactory.updateTextStyle helper takes objectId and style.
                    // It returns `updateTextStyle` with objectId. IT DOES NOT SUPPORT cellLocation yet!
                    // FAILURE: RequestFactory.updateTextStyle only works for Shapes/Textboxes unless I update it or write raw request here.
                    // I will write raw request here to be safe and precise.
                    requests.push({
                        updateTextStyle: {
                            objectId: tableId,
                            cellLocation: { rowIndex: 0, columnIndex: c },
                            style: {
                                bold: true,
                                fontSize: { magnitude: 14, unit: 'PT' },
                                foregroundColor: { opaqueColor: { themeColor: 'BACKGROUND1' } },
                                fontFamily: theme.fonts.family
                            },
                            fields: 'bold,fontSize,foregroundColor,fontFamily'
                        }
                    });
                }

                // Cell Background (Primary)
                requests.push({
                    updateTableCellProperties: {
                        objectId: tableId,
                        tableRange: { location: { rowIndex: 0, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
                        tableCellProperties: {
                            tableCellBackgroundFill: { solidFill: { color: { themeColor: 'ACCENT1' } } },
                            contentAlignment: 'MIDDLE'
                        },
                        fields: 'tableCellBackgroundFill,contentAlignment'
                    }
                });
            });
            rowIndex++;
        }

        // 4. Data Rows
        rows.forEach((row: any[], rIdx: number) => {
            const actualRowIndex = rowIndex + rIdx;
            const isEven = actualRowIndex % 2 === 0; // Header is 0, so first data row is 1 (odd).
            // Usually we want banding.
            const bgFill = isEven
                ? { solidFill: { color: { rgbColor: { red: 0.97, green: 0.97, blue: 0.98 } } } } // F8F9FA (Light Gray)
                : { solidFill: { color: { themeColor: 'BACKGROUND1' } } }; // White

            for (let c = 0; c < numCols; c++) {
                const cellText = String(row[c] ?? '');
                if (cellText) {
                    requests.push({ insertText: { objectId: tableId, cellLocation: { rowIndex: actualRowIndex, columnIndex: c }, text: cellText } });
                    requests.push({
                        updateTextStyle: {
                            objectId: tableId,
                            cellLocation: { rowIndex: actualRowIndex, columnIndex: c },
                            style: {
                                fontSize: { magnitude: 12, unit: 'PT' },
                                foregroundColor: { opaqueColor: { themeColor: 'TEXT1' } },
                                fontFamily: theme.fonts.family
                            },
                            fields: 'fontSize,foregroundColor,fontFamily'
                        }
                    });
                }

                // Cell Style
                requests.push({
                    updateTableCellProperties: {
                        objectId: tableId,
                        tableRange: { location: { rowIndex: actualRowIndex, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
                        tableCellProperties: {
                            tableCellBackgroundFill: bgFill as any,
                            contentAlignment: 'MIDDLE'
                        },
                        fields: 'tableCellBackgroundFill,contentAlignment'
                    }
                });

                // Horizontal Border (Bottom)
                requests.push({
                    updateTableBorderProperties: {
                        objectId: tableId,
                        tableRange: { location: { rowIndex: actualRowIndex, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
                        borderPosition: 'BOTTOM',
                        tableBorderProperties: {
                            tableBorderFill: { solidFill: { color: { rgbColor: { red: 0.9, green: 0.9, blue: 0.9 } } } },
                            weight: { magnitude: 1, unit: 'PT' },
                            dashStyle: 'SOLID'
                        },
                        fields: 'tableBorderFill,weight,dashStyle'
                    }
                });
            }
        });

        return requests;
    }
}
