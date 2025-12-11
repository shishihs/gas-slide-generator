import { IDiagramRenderer } from './IDiagramRenderer';
import { CardsDiagramRenderer } from './CardsDiagramRenderer';

export class DiagramRendererFactory {
    static getRenderer(type: string): IDiagramRenderer | null {
        const normalizedType = type.toLowerCase();

        if (normalizedType.includes('cards') || normalizedType.includes('headercards')) {
            return new CardsDiagramRenderer();
        }

        return null;
    }
}
