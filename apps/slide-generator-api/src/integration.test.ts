import { vi, describe, it, expect } from 'vitest';
import { PresentationApplicationService } from '../src/application/PresentationApplicationService';
import { GasSlideRepository } from '../src/infrastructure/gas/GasSlideRepository';
import { setupGasGlobals } from '../src/test/gas-mocks';

let slidesApp: any;
let driveApp: any;

beforeEach(() => {
    const mocks = setupGasGlobals();
    slidesApp = mocks.slidesApp;
    driveApp = mocks.driveApp;

    // Override specific behavior for integration test
    slidesApp.create.mockImplementation((title) => ({
        getUrl: vi.fn().mockReturnValue('https://mock-slide.com'),
        getPageWidth: vi.fn().mockReturnValue(960),
        getPageHeight: vi.fn().mockReturnValue(540),
        getSlides: vi.fn().mockReturnValue([]),
        getMasters: vi.fn().mockReturnValue([]),
        getLayouts: vi.fn().mockReturnValue([
            { getLayoutName: vi.fn().mockReturnValue('TITLE') },
            { getLayoutName: vi.fn().mockReturnValue('SECTION_HEADER') },
            { getLayoutName: vi.fn().mockReturnValue('TITLE_AND_BODY') }
        ]),
        appendSlide: vi.fn().mockImplementation(() => ({
            getBackground: vi.fn().mockReturnValue({
                setSolidFill: vi.fn()
            }),
            insertShape: vi.fn().mockReturnValue({
                getText: vi.fn().mockReturnValue({
                    setText: vi.fn(),
                    getTextStyle: vi.fn().mockReturnValue({
                        setFontFamily: vi.fn().mockReturnThis(),
                        setFontSize: vi.fn().mockReturnThis(),
                        setBold: vi.fn().mockReturnThis(),
                        setForegroundColor: vi.fn().mockReturnThis()
                    }),
                    getParagraphs: vi.fn().mockReturnValue([])
                }),
                getBorder: vi.fn().mockReturnValue({
                    setTransparent: vi.fn()
                }),
                getFill: vi.fn().mockReturnValue({
                    setSolidFill: vi.fn()
                }),
                setContentAlignment: vi.fn()
            }),
            getPlaceholder: vi.fn().mockReturnValue({
                asShape: vi.fn().mockReturnValue({
                    getText: vi.fn().mockReturnValue({
                        setText: vi.fn(),
                        getParagraphs: vi.fn().mockReturnValue([]),
                        getTextStyle: vi.fn().mockReturnValue({
                            setFontFamily: vi.fn().mockReturnThis(),
                            setFontSize: vi.fn().mockReturnThis(),
                            setBold: vi.fn().mockReturnThis(),
                            setForegroundColor: vi.fn().mockReturnThis()
                        })
                    })
                })
            }),
            getPlaceholders: vi.fn().mockReturnValue([{
                getPlaceholderType: vi.fn().mockReturnValue('BODY'),
                asShape: vi.fn().mockReturnValue({
                    getText: vi.fn().mockReturnValue({
                        setText: vi.fn(),
                        getParagraphs: vi.fn().mockReturnValue([]),
                        getTextStyle: vi.fn().mockReturnValue({
                            setFontFamily: vi.fn().mockReturnThis(),
                            setFontSize: vi.fn().mockReturnThis(),
                            setBold: vi.fn().mockReturnThis(),
                            setForegroundColor: vi.fn().mockReturnThis()
                        })
                    }),
                    getLeft: vi.fn().mockReturnValue(100)
                })
            }]),
            getNotesPage: vi.fn().mockReturnValue({
                getSpeakerNotesShape: vi.fn().mockReturnValue({
                    getText: vi.fn().mockReturnValue({ setText: vi.fn() })
                })
            })
        }))
    }));

    slidesApp.openById.mockImplementation(() => ({
        getUrl: vi.fn().mockReturnValue('https://template-copy.com'),
        getPageWidth: vi.fn().mockReturnValue(960),
        getPageHeight: vi.fn().mockReturnValue(540),
        getSlides: vi.fn().mockReturnValue([{ remove: vi.fn() }, { remove: vi.fn() }]),
        getMasters: vi.fn().mockReturnValue([]),
        getLayouts: vi.fn().mockReturnValue([
            { getLayoutName: vi.fn().mockReturnValue('TITLE') },
            { getLayoutName: vi.fn().mockReturnValue('SECTION_HEADER') },
            { getLayoutName: vi.fn().mockReturnValue('TITLE_AND_BODY') }
        ]),
        appendSlide: vi.fn().mockImplementation(() => ({
            getBackground: vi.fn().mockReturnValue({
                setSolidFill: vi.fn()
            }),
            insertShape: vi.fn().mockReturnValue({
                getText: vi.fn().mockReturnValue({
                    setText: vi.fn(),
                    getTextStyle: vi.fn().mockReturnValue({
                        setFontFamily: vi.fn().mockReturnThis(),
                        setFontSize: vi.fn().mockReturnThis(),
                        setBold: vi.fn().mockReturnThis(),
                        setForegroundColor: vi.fn().mockReturnThis()
                    }),
                    getParagraphs: vi.fn().mockReturnValue([])
                }),
                getBorder: vi.fn().mockReturnValue({
                    setTransparent: vi.fn()
                }),
                getFill: vi.fn().mockReturnValue({
                    setSolidFill: vi.fn()
                }),
                setContentAlignment: vi.fn()
            }),
            getPlaceholder: vi.fn().mockReturnValue({
                asShape: vi.fn().mockReturnValue({
                    getText: vi.fn().mockReturnValue({
                        setText: vi.fn(),
                        getParagraphs: vi.fn().mockReturnValue([]),
                        getTextStyle: vi.fn().mockReturnValue({
                            setFontFamily: vi.fn().mockReturnThis(),
                            setFontSize: vi.fn().mockReturnThis(),
                            setBold: vi.fn().mockReturnThis(),
                            setForegroundColor: vi.fn().mockReturnThis()
                        })
                    })
                })
            }),
            getPlaceholders: vi.fn().mockReturnValue([{
                getPlaceholderType: vi.fn().mockReturnValue('BODY'),
                asShape: vi.fn().mockReturnValue({
                    getText: vi.fn().mockReturnValue({
                        setText: vi.fn(),
                        getParagraphs: vi.fn().mockReturnValue([]),
                        getTextStyle: vi.fn().mockReturnValue({
                            setFontFamily: vi.fn().mockReturnThis(),
                            setFontSize: vi.fn().mockReturnThis(),
                            setBold: vi.fn().mockReturnThis(),
                            setForegroundColor: vi.fn().mockReturnThis()
                        })
                    }),
                    getLeft: vi.fn().mockReturnValue(100)
                })
            }]),
            getNotesPage: vi.fn().mockReturnValue({
                getSpeakerNotesShape: vi.fn().mockReturnValue({
                    getText: vi.fn().mockReturnValue({ setText: vi.fn() })
                })
            })
        }))
    }));

    driveApp.getFileById.mockReturnValue({
        makeCopy: vi.fn().mockReturnValue({
            getId: vi.fn().mockReturnValue('new-copy-id')
        }),
        getBlob: vi.fn()
    });
});

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
