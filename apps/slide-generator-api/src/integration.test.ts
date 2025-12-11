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
        getId: vi.fn().mockReturnValue('mock-presentation-id'),
        getUrl: vi.fn().mockReturnValue('https://mock-slide.com'),
        getPageWidth: vi.fn().mockReturnValue(960),
        getPageHeight: vi.fn().mockReturnValue(540),
        getSlides: vi.fn().mockReturnValue([]),
        getMasters: vi.fn().mockReturnValue([]),
        getLayouts: vi.fn().mockReturnValue([
            { getLayoutName: vi.fn().mockReturnValue('TITLE'), getObjectId: vi.fn().mockReturnValue('layout-title-id') },
            { getLayoutName: vi.fn().mockReturnValue('SECTION_HEADER'), getObjectId: vi.fn().mockReturnValue('layout-section-id') },
            { getLayoutName: vi.fn().mockReturnValue('TITLE_AND_BODY'), getObjectId: vi.fn().mockReturnValue('layout-content-id') }
        ]),
        appendSlide: vi.fn().mockImplementation(() => ({
            getObjectId: vi.fn().mockReturnValue('mock-slide-id'),
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
        getId: vi.fn().mockReturnValue('mock-presentation-id'),
        getUrl: vi.fn().mockReturnValue('https://template-copy.com'),
        getPageWidth: vi.fn().mockReturnValue(960),
        getPageHeight: vi.fn().mockReturnValue(540),
        getSlides: vi.fn().mockReturnValue([{ remove: vi.fn() }, { remove: vi.fn() }]),
        getMasters: vi.fn().mockReturnValue([]),
        getLayouts: vi.fn().mockReturnValue([
            { getLayoutName: vi.fn().mockReturnValue('TITLE'), getObjectId: vi.fn().mockReturnValue('layout-title-id') },
            { getLayoutName: vi.fn().mockReturnValue('SECTION_HEADER'), getObjectId: vi.fn().mockReturnValue('layout-section-id') },
            { getLayoutName: vi.fn().mockReturnValue('TITLE_AND_BODY'), getObjectId: vi.fn().mockReturnValue('layout-content-id') },
            { getLayoutName: vi.fn().mockReturnValue('BLANK'), getObjectId: vi.fn().mockReturnValue('layout-blank-id') }
        ]),
        appendSlide: vi.fn().mockImplementation(() => ({
            getObjectId: vi.fn().mockReturnValue('mock-slide-id-2'),
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

    it('should handle diagram slides without error', () => {
        // Reproduce 'insertShape' error by requesting a diagram
        const repo = new GasSlideRepository();
        const service = new PresentationApplicationService(repo);

        // Assume layout map handles mapping or fallback to CONTENT
        // The error happens inside Generator -> Renderer
        service.createPresentation({
            title: 'Diagram Pres',
            slides: [
                {
                    title: 'Triangle Diagram',
                    // @ts-ignore
                    layout: 'TIMELINE', // Request something that triggers diagram generator
                    rawData: { type: 'triangle', items: [{ title: 'A' }, { title: 'B' }] },
                    content: []
                }
            ]
        });

        // If successful, batchUpdate should be called.
        // If it fails with 'slide.insertShape is not a function', this test will fail.
        expect((global as any).Slides.Presentations.batchUpdate).toHaveBeenCalled();
    });
});
