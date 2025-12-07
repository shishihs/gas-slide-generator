import { Slide } from '../../src/domain/model/Slide';
import { SlideContent, SlideTitle } from '../../src/domain/model/SlideElement';

describe('Slide Entity', () => {
    it('should be created with a title and content', () => {
        const title = new SlideTitle('Test Title');
        const content = new SlideContent(['Point 1', 'Point 2']);
        const slide = new Slide(title, content);

        expect(slide.title.value).toBe('Test Title');
        expect(slide.content.items).toHaveLength(2);
        expect(slide.content.items[0]).toBe('Point 1');
    });

    it('should allow optional speaker notes', () => {
        const title = new SlideTitle('Title');
        const content = new SlideContent([]);
        const slide = new Slide(title, content, 'Notes here');

        expect(slide.notes).toBe('Notes here');
    });
});
