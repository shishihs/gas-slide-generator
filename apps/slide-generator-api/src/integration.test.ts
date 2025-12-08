import { PresentationApplicationService } from '../src/application/PresentationApplicationService';
import { GasSlideRepository } from '../src/infrastructure/gas/GasSlideRepository';
import { Presentation } from '../src/domain/model/Presentation';

const mockSlidesApp = {
    create: jest.fn().mockImplementation((title) => ({
        getUrl: jest.fn().mockReturnValue('https://mock-slide.com'),
        getPageWidth: jest.fn().mockReturnValue(960),
        getPageHeight: jest.fn().mockReturnValue(540),
        getSlides: jest.fn().mockReturnValue([]),
        getMasters: jest.fn().mockReturnValue([]),
        getLayouts: jest.fn().mockReturnValue([
            { getLayoutName: jest.fn().mockReturnValue('TITLE') },
            { getLayoutName: jest.fn().mockReturnValue('SECTION_HEADER') },
            { getLayoutName: jest.fn().mockReturnValue('TITLE_AND_BODY') }
        ]),
        appendSlide: jest.fn().mockImplementation(() => ({
            getBackground: jest.fn().mockReturnValue({
                setSolidFill: jest.fn()
            }),
            insertShape: jest.fn().mockReturnValue({
                getText: jest.fn().mockReturnValue({
                    setText: jest.fn(),
                    getTextStyle: jest.fn().mockReturnValue({
                        setFontFamily: jest.fn().mockReturnThis(),
                        setFontSize: jest.fn().mockReturnThis(),
                        setBold: jest.fn().mockReturnThis(),
                        setForegroundColor: jest.fn().mockReturnThis()
                    }),
                    getParagraphs: jest.fn().mockReturnValue([])
                }),
                getBorder: jest.fn().mockReturnValue({
                    setTransparent: jest.fn()
                }),
                getFill: jest.fn().mockReturnValue({
                    setSolidFill: jest.fn()
                }),
                setContentAlignment: jest.fn()
            }),
            getPlaceholder: jest.fn().mockReturnValue({
                asShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({
                        setText: jest.fn(),
                        getParagraphs: jest.fn().mockReturnValue([]),
                        getTextStyle: jest.fn().mockReturnValue({
                            setFontFamily: jest.fn().mockReturnThis(),
                            setFontSize: jest.fn().mockReturnThis(),
                            setBold: jest.fn().mockReturnThis(),
                            setForegroundColor: jest.fn().mockReturnThis()
                        })
                    })
                })
            }),
            getPlaceholders: jest.fn().mockReturnValue([{
                getPlaceholderType: jest.fn().mockReturnValue('BODY'),
                asShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({
                        setText: jest.fn(),
                        getParagraphs: jest.fn().mockReturnValue([]),
                        getTextStyle: jest.fn().mockReturnValue({
                            setFontFamily: jest.fn().mockReturnThis(),
                            setFontSize: jest.fn().mockReturnThis(),
                            setBold: jest.fn().mockReturnThis(),
                            setForegroundColor: jest.fn().mockReturnThis()
                        })
                    }),
                    getLeft: jest.fn().mockReturnValue(100)
                })
            }]),
            getNotesPage: jest.fn().mockReturnValue({
                getSpeakerNotesShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({ setText: jest.fn() })
                })
            })
        }))
    })),
    openById: jest.fn().mockImplementation(() => ({
        getUrl: jest.fn().mockReturnValue('https://template-copy.com'),
        getPageWidth: jest.fn().mockReturnValue(960),
        getPageHeight: jest.fn().mockReturnValue(540),
        getSlides: jest.fn().mockReturnValue([{ remove: jest.fn() }, { remove: jest.fn() }]),
        getMasters: jest.fn().mockReturnValue([]),
        getLayouts: jest.fn().mockReturnValue([
            { getLayoutName: jest.fn().mockReturnValue('TITLE') },
            { getLayoutName: jest.fn().mockReturnValue('SECTION_HEADER') },
            { getLayoutName: jest.fn().mockReturnValue('TITLE_AND_BODY') }
        ]),
        appendSlide: jest.fn().mockImplementation(() => ({
            getBackground: jest.fn().mockReturnValue({
                setSolidFill: jest.fn()
            }),
            insertShape: jest.fn().mockReturnValue({
                getText: jest.fn().mockReturnValue({
                    setText: jest.fn(),
                    getTextStyle: jest.fn().mockReturnValue({
                        setFontFamily: jest.fn().mockReturnThis(),
                        setFontSize: jest.fn().mockReturnThis(),
                        setBold: jest.fn().mockReturnThis(),
                        setForegroundColor: jest.fn().mockReturnThis()
                    }),
                    getParagraphs: jest.fn().mockReturnValue([])
                }),
                getBorder: jest.fn().mockReturnValue({
                    setTransparent: jest.fn()
                }),
                getFill: jest.fn().mockReturnValue({
                    setSolidFill: jest.fn()
                }),
                setContentAlignment: jest.fn()
            }),
            getPlaceholder: jest.fn().mockReturnValue({
                asShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({
                        setText: jest.fn(),
                        getParagraphs: jest.fn().mockReturnValue([]),
                        getTextStyle: jest.fn().mockReturnValue({
                            setFontFamily: jest.fn().mockReturnThis(),
                            setFontSize: jest.fn().mockReturnThis(),
                            setBold: jest.fn().mockReturnThis(),
                            setForegroundColor: jest.fn().mockReturnThis()
                        })
                    })
                })
            }),
            getPlaceholders: jest.fn().mockReturnValue([{
                getPlaceholderType: jest.fn().mockReturnValue('BODY'),
                asShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({
                        setText: jest.fn(),
                        getParagraphs: jest.fn().mockReturnValue([]),
                        getTextStyle: jest.fn().mockReturnValue({
                            setFontFamily: jest.fn().mockReturnThis(),
                            setFontSize: jest.fn().mockReturnThis(),
                            setBold: jest.fn().mockReturnThis(),
                            setForegroundColor: jest.fn().mockReturnThis()
                        })
                    }),
                    getLeft: jest.fn().mockReturnValue(100)
                })
            }]),
            getNotesPage: jest.fn().mockReturnValue({
                getSpeakerNotesShape: jest.fn().mockReturnValue({
                    getText: jest.fn().mockReturnValue({ setText: jest.fn() })
                })
            })
        }))
    })),
    PredefinedLayout: {
        TITLE_AND_BODY: 'TITLE_AND_BODY',
        TITLE: 'TITLE',
        SECTION_HEADER: 'SECTION_HEADER',
        BLANK: 'BLANK' // Added BLANK
    },
    PlaceholderType: {
        TITLE: 'TITLE',
        CENTERED_TITLE: 'CENTERED_TITLE',
        SUBTITLE: 'SUBTITLE',
        BODY: 'BODY'
    },
    ShapeType: {
        TEXT_BOX: 'TEXT_BOX',
        RECTANGLE: 'RECTANGLE',
        ELLIPSE: 'ELLIPSE'
    },
    ArrowStyle: {
        FILL_ARROW: 'FILL_ARROW'
    },
    LineCategory: {
        STRAIGHT: 'STRAIGHT'
    },
    ParagraphAlignment: {
        START: 'START',
        CENTER: 'CENTER',
        END: 'END',
        JUSTIFIED: 'JUSTIFIED'
    },
    ContentAlignment: {
        TOP: 'TOP',
        MIDDLE: 'MIDDLE',
        BOTTOM: 'BOTTOM'
    }
};

(global as any).SlidesApp = mockSlidesApp;

(global as any).Logger = {
    log: jest.fn()
};

(global as any).DriveApp = {
    getFileById: jest.fn().mockReturnValue({
        makeCopy: jest.fn().mockReturnValue({
            getId: jest.fn().mockReturnValue('new-copy-id')
        }),
        getBlob: jest.fn()
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
