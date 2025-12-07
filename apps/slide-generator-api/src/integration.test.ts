import { PresentationApplicationService } from '../src/application/PresentationApplicationService';
import { GasSlideRepository } from '../src/infrastructure/gas/GasSlideRepository';
import { Presentation } from '../src/domain/model/Presentation';

// Mock GAS globals
(global as any).SlidesApp = {
    create: jest.fn().mockReturnValue({
        getUrl: jest.fn().mockReturnValue('https://mock-slide.com'),
        getSlides: jest.fn().mockReturnValue([]),
        getMasters: jest.fn().mockReturnValue([]),
        getLayouts: jest.fn().mockReturnValue([]),
        appendSlide: jest.fn().mockReturnValue({
            getPlaceholder: jest.fn().mockReturnValue({
                asShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({
                        setText: jest.fn(),
                        getParagraphs: jest.fn().mockReturnValue([])
                    })
                })
            }),
            getNotesPage: jest.fn().mockReturnValue({
                getSpeakerNotesShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({ setText: jest.fn() })
                })
            })
        }),
        PredefinedLayout: { TITLE_AND_BODY: 'TITLE_AND_BODY', TITLE: 'TITLE' },
        PlaceholderType: {
            TITLE: 'TITLE',
            CENTERED_TITLE: 'CENTERED_TITLE',
            SUBTITLE: 'SUBTITLE',
            BODY: 'BODY'
        }
    }),
    openById: jest.fn().mockReturnValue({
        getUrl: jest.fn().mockReturnValue('https://template-copy.com'),
        getSlides: jest.fn().mockReturnValue([{}, {}]), // 2 existing slides
        getMasters: jest.fn().mockReturnValue([]),
        getLayouts: jest.fn().mockReturnValue([]),
        appendSlide: jest.fn().mockReturnValue({
            getPlaceholder: jest.fn().mockReturnValue({
                asShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({
                        setText: jest.fn(),
                        getParagraphs: jest.fn().mockReturnValue([])
                    })
                })
            }),
            getNotesPage: jest.fn().mockReturnValue({
                getSpeakerNotesShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({ setText: jest.fn() })
                })
            })
        })
    })
};

(global as any).DriveApp = {
    getFileById: jest.fn().mockReturnValue({
        makeCopy: jest.fn().mockReturnValue({
            getId: jest.fn().mockReturnValue('new-copy-id')
        })
    })
};

describe('Integration: Application Service + Infrastructure', () => {
    it('should create blank presentation if no templateId provided', () => {
        const repo = new GasSlideRepository();
        const service = new PresentationApplicationService(repo);

        const url = service.createPresentation({
            title: 'Blank Pres',
            slides: [{ title: 'Intro', content: ['A'] }]
        });

        expect(url).toBe('https://mock-slide.com');
        expect((global as any).SlidesApp.create).toHaveBeenCalledWith('Blank Pres');
    });

    it('should copy template if templateId provided', () => {
        const repo = new GasSlideRepository();
        const service = new PresentationApplicationService(repo);

        const url = service.createPresentation({
            title: 'Templated Pres',
            templateId: 'tmpl-abc',
            slides: [{ title: 'Intro', content: ['A'] }]
        });

        expect(url).toBe('https://template-copy.com');
        expect((global as any).DriveApp.getFileById).toHaveBeenCalledWith('tmpl-abc');
        expect((global as any).SlidesApp.openById).toHaveBeenCalledWith('new-copy-id');
    });
});
