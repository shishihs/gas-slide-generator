import { ISlideRepository } from '../../domain/repositories/ISlideRepository';
import { Presentation } from '../../domain/model/Presentation';

// This class acts as an Anti-Corruption Layer (ACL) adaptation
// It translates Domain objects into GAS API calls.
export class GasSlideRepository implements ISlideRepository {

    createPresentation(presentation: Presentation): string {
        const slidesApp = SlidesApp; // Global access in GAS environment
        const pres = slidesApp.create(presentation.title);

        // Remove the default first slide if desired, or use it
        const slides = pres.getSlides();
        if (slides.length > 0) {
            slides[0].remove();
        }

        // Apply basic Logic (This will be expanded to use the legacy chart logic later)
        presentation.slides.forEach(slideModel => {
            const slide = pres.appendSlide(slidesApp.PredefinedLayout.TITLE_AND_BODY);

            // Set Title
            const titleShape = slide.getPlaceholder(slidesApp.PlaceholderType.TITLE);
            if (titleShape) {
                titleShape.asShape().getText().setText(slideModel.title.value);
            }

            // Set Body (Bullets)
            const bodyShape = slide.getPlaceholder(slidesApp.PlaceholderType.BODY);
            if (bodyShape) {
                const textRange = bodyShape.asShape().getText();
                textRange.setText(slideModel.content.items.join('\n'));

                // Basic Bullet points
                const paragraphs = textRange.getParagraphs();
                for (let i = 0; i < paragraphs.length; i++) {
                    paragraphs[i].setBullet(true);
                }
            }

            // Set Notes
            if (slideModel.notes) {
                const notesPage = slide.getNotesPage();
                notesPage.getSpeakerNotesShape().getText().setText(slideModel.notes);
            }
        });

        return pres.getUrl();
    }
}
