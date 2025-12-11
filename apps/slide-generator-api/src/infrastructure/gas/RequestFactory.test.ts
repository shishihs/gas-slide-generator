import { describe, it, expect } from 'vitest';
import { RequestFactory } from './RequestFactory';

describe('RequestFactory', () => {
    describe('updateShapeProperties', () => {
        it('should handle valid hex colors', () => {
            const req = RequestFactory.updateShapeProperties('id', '#FF0000', '#0000FF');
            expect(req).toBeDefined();
            // @ts-ignore
            const fill = req.updateShapeProperties.shapeProperties.shapeBackgroundFill.solidFill.color.rgbColor;
            expect(fill.red).toBeCloseTo(1);
            expect(fill.green).toBe(0);
            expect(fill.blue).toBe(0);
        });

        it('should safely handle null fillHex', () => {
            // @ts-ignore
            const req = RequestFactory.updateShapeProperties('id', null, null);
            expect(req).toBeDefined();
            // Should not set shapeBackgroundFill if null
            // @ts-ignore
            expect(req.updateShapeProperties.shapeProperties.shapeBackgroundFill).toBeUndefined();
        });

        it('should safely handle undefined fillHex', () => {
            const req = RequestFactory.updateShapeProperties('id', undefined, undefined);
            expect(req).toBeDefined();
            // @ts-ignore
            expect(req.updateShapeProperties.shapeProperties.shapeBackgroundFill).toBeUndefined();
        });

        it('should safely handle short/invalid hex', () => {
            // @ts-ignore
            const req = RequestFactory.updateShapeProperties('id', '#123', 'invalid');
            expect(req).toBeDefined();
            // Should ignore invalid hex formats to avoid substring errors
            // @ts-ignore
            expect(req.updateShapeProperties.shapeProperties.shapeBackgroundFill).toBeUndefined();
        });

        it('should handle TRANSPARENT', () => {
            const req = RequestFactory.updateShapeProperties('id', 'TRANSPARENT', 'TRANSPARENT');
            // @ts-ignore
            expect(req.updateShapeProperties.shapeProperties.shapeBackgroundFill.propertyState).toBe('NOT_RENDERED');
            // @ts-ignore
            expect(req.updateShapeProperties.shapeProperties.outline.propertyState).toBe('NOT_RENDERED');
        });
    });

    describe('updateTextStyle', () => {
        it('should handle valid hex color', () => {
            const req = RequestFactory.updateTextStyle('id', { color: '#00FF00' });
            // @ts-ignore
            const color = req.updateTextStyle.style.foregroundColor.opaqueColor.rgbColor;
            expect(color.green).toBeCloseTo(1);
        });

        it('should safely handle short hex color', () => {
            // Should not throw
            const req = RequestFactory.updateTextStyle('id', { color: '#1' });
            // @ts-ignore
            // It defaults to DARK1 if it starts with # but isn't a valid hex for RGB conversion
            expect(req.updateTextStyle.style.foregroundColor.opaqueColor.themeColor).toBe('DARK1');
        });
    });
});
