
export class RequestFactory {
    static createSlide(pageId: string, layoutId: string = 'TITLE_AND_BODY', isPredefined: boolean = true): GoogleAppsScript.Slides.Schema.Request {
        return {
            createSlide: {
                objectId: pageId,
                slideLayoutReference: isPredefined ? { predefinedLayout: layoutId } : { layoutId: layoutId }
            }
        };
    }

    static createTextBox(pageId: string, objectId: string, text: string, leftOpt: number, topOpt: number, widthOpt: number, heightOpt: number): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];

        // Create Shape
        requests.push({
            createShape: {
                objectId: objectId,
                shapeType: 'TEXT_BOX',
                elementProperties: {
                    pageObjectId: pageId,
                    size: {
                        width: { magnitude: widthOpt, unit: 'PT' },
                        height: { magnitude: heightOpt, unit: 'PT' }
                    },
                    transform: {
                        scaleX: 1,
                        scaleY: 1,
                        translateX: leftOpt,
                        translateY: topOpt,
                        unit: 'PT'
                    }
                }
            }
        });

        // Insert Text
        if (text) {
            requests.push({
                insertText: {
                    objectId: objectId,
                    text: text
                }
            });
        }

        return requests;
    }

    static updateTextStyle(objectId: string, style: { bold?: boolean, fontSize?: number, color?: string, fontFamily?: string }, selectionStart?: number, selectionEnd?: number, rowIndex?: number, colIndex?: number): GoogleAppsScript.Slides.Schema.Request {
        const fields: string[] = [];
        const textStyle: any = {};

        if (style.bold !== undefined) {
            textStyle.bold = style.bold;
            fields.push('bold');
        }
        if (style.fontSize !== undefined) {
            textStyle.fontSize = { magnitude: style.fontSize, unit: 'PT' };
            fields.push('fontSize');
        }
        if (style.color !== undefined) {
            textStyle.foregroundColor = {
                opaqueColor: { themeColor: style.color.startsWith('#') ? 'DARK1' : style.color as any }
            };
            const rgb = RequestFactory.toRgbColor(style.color);
            if (rgb) {
                textStyle.foregroundColor = {
                    opaqueColor: { rgbColor: rgb }
                };
            }
            fields.push('foregroundColor');
        }
        if (style.fontFamily !== undefined) {
            textStyle.fontFamily = style.fontFamily;
            fields.push('fontFamily');
        }

        const request: any = {
            updateTextStyle: {
                objectId: objectId,
                style: textStyle,
                fields: fields.join(',')
            }
        };

        if (selectionStart !== undefined && selectionEnd !== undefined) {
            request.updateTextStyle.textRange = {
                type: 'FIXED_RANGE',
                startIndex: selectionStart,
                endIndex: selectionStart + selectionEnd
            };
        }

        if (rowIndex !== undefined && colIndex !== undefined) {
            request.updateTextStyle.cellLocation = {
                rowIndex: rowIndex,
                columnIndex: colIndex
            };
        }

        return request;
    }

    static createShape(pageId: string, objectId: string, shapeType: string, left: number, top: number, width: number, height: number): GoogleAppsScript.Slides.Schema.Request {
        return {
            createShape: {
                objectId: objectId,
                shapeType: shapeType as any,
                elementProperties: {
                    pageObjectId: pageId,
                    size: {
                        width: { magnitude: width, unit: 'PT' },
                        height: { magnitude: height, unit: 'PT' }
                    },
                    transform: {
                        scaleX: 1,
                        scaleY: 1,
                        translateX: left,
                        translateY: top,
                        unit: 'PT'
                    }
                }
            }
        };
    }

    static createLine(pageId: string, objectId: string, x1: number, y1: number, x2: number, y2: number): GoogleAppsScript.Slides.Schema.Request {
        // Create line category straight
        // Note: API createLine arguments are elementProperties (transform + size) and lineCategory
        // Calculating transform/size from coords is tricky for lines in API.
        // Actually createLine takes elementProperties.
        // Simpler to just create it at 0,0 and then transform? 
        // Or strictly calculated:
        // Width = x2-x1, Height = y2-y1. 
        const width = Math.abs(x2 - x1) || 1; // avoid 0
        const height = Math.abs(y2 - y1) || 1;
        const left = Math.min(x1, x2);
        const top = Math.min(y1, y2);

        return {
            createLine: {
                objectId: objectId,
                lineCategory: 'STRAIGHT',
                elementProperties: {
                    pageObjectId: pageId,
                    size: {
                        width: { magnitude: width, unit: 'PT' },
                        height: { magnitude: height, unit: 'PT' }
                    },
                    transform: {
                        scaleX: 1,
                        scaleY: 1,
                        translateX: left,
                        translateY: top,
                        unit: 'PT'
                    }
                }
            }
        };
        // Note: Directionality might be lost (always top-left to bottom-right visual box).
        // API doesn't specify start/end points easily like SlidesApp. 
        // For now this is a basic approximation.
    }

    static updateShapeProperties(objectId: string, fillHex?: string, borderHex?: string, borderWeightPt?: number, contentAlignment?: string): GoogleAppsScript.Slides.Schema.Request {
        const fields: string[] = [];
        const shapeProperties: any = {};

        if (fillHex !== undefined && fillHex !== null) {
            // Check for transparent
            if (fillHex === 'TRANSPARENT') {
                shapeProperties.shapeBackgroundFill = { propertyState: 'NOT_RENDERED' };
            } else {
                const rgb = RequestFactory.toRgbColor(fillHex);
                if (rgb) {
                    shapeProperties.shapeBackgroundFill = {
                        solidFill: { color: { rgbColor: rgb } }
                    };
                }
            }
            fields.push('shapeBackgroundFill');
        }

        if (borderHex !== undefined || borderWeightPt !== undefined) {
            const outline: any = {};
            if (borderHex === 'TRANSPARENT') {
                outline.propertyState = 'NOT_RENDERED';
            } else {
                const rgb = RequestFactory.toRgbColor(borderHex);
                if (rgb) {
                    outline.outlineFill = { solidFill: { color: { rgbColor: rgb } } };
                }
            }
            if (borderWeightPt !== undefined) {
                outline.weight = { magnitude: borderWeightPt, unit: 'PT' };
            }
            shapeProperties.outline = outline;
            fields.push('outline');
        }

        if (contentAlignment) {
            shapeProperties.contentAlignment = contentAlignment;
            fields.push('contentAlignment');
        }

        return {
            updateShapeProperties: {
                objectId: objectId,
                shapeProperties: shapeProperties,
                fields: fields.join(',')
            }
        };
    }

    static updateParagraphStyle(objectId: string, alignment: string, rowIndex?: number, colIndex?: number): GoogleAppsScript.Slides.Schema.Request {
        const req: any = {
            updateParagraphStyle: {
                objectId: objectId,
                style: {
                    alignment: alignment as any
                },
                fields: 'alignment'
            }
        };
        if (rowIndex !== undefined && colIndex !== undefined) {
            req.updateParagraphStyle.cellLocation = {
                rowIndex: rowIndex,
                columnIndex: colIndex
            };
        }
        return req;
    }

    static createTable(pageId: string, objectId: string, rows: number, cols: number, left: number, top: number, width: number, height: number): GoogleAppsScript.Slides.Schema.Request {
        return {
            createTable: {
                objectId: objectId,
                rows: rows,
                columns: cols,
                elementProperties: {
                    pageObjectId: pageId,
                    size: {
                        width: { magnitude: width, unit: 'PT' },
                        height: { magnitude: height, unit: 'PT' }
                    },
                    transform: {
                        scaleX: 1,
                        scaleY: 1,
                        translateX: left,
                        translateY: top,
                        unit: 'PT'
                    }
                }
            }
        };
    }

    static insertText(objectId: string, text: string, rowIndex?: number, colIndex?: number): GoogleAppsScript.Slides.Schema.Request {
        const req: any = {
            insertText: {
                objectId: objectId,
                text: text
            }
        };
        if (rowIndex !== undefined && colIndex !== undefined) {
            req.insertText.cellLocation = {
                rowIndex: rowIndex,
                columnIndex: colIndex
            };
        }
        return req;
    }

    static createImage(pageId: string, objectId: string, url: string, left: number, top: number, width: number, height: number): GoogleAppsScript.Slides.Schema.Request {
        return {
            createImage: {
                objectId: objectId,
                url: url,
                elementProperties: {
                    pageObjectId: pageId,
                    size: {
                        width: { magnitude: width, unit: 'PT' },
                        height: { magnitude: height, unit: 'PT' }
                    },
                    transform: {
                        scaleX: 1,
                        scaleY: 1,
                        translateX: left,
                        translateY: top,
                        unit: 'PT'
                    }
                }
            }
        };
    }

    static toRgbColor(hex: string): GoogleAppsScript.Slides.Schema.RgbColor | undefined {
        if (!hex || typeof hex !== 'string' || !hex.startsWith('#') || hex.length < 7) {
            return undefined;
        }
        const r = parseInt(hex.substring(1, 3), 16) / 255;
        const g = parseInt(hex.substring(3, 5), 16) / 255;
        const b = parseInt(hex.substring(5, 7), 16) / 255;
        return { red: r, green: g, blue: b };
    }

    static groupObjects(pageId: string, groupObjectId: string, childrenObjectIds: string[]): GoogleAppsScript.Slides.Schema.Request {
        return {
            groupObjects: {
                groupObjectId: groupObjectId,
                childrenObjectIds: childrenObjectIds
            }
        };
    }
}
