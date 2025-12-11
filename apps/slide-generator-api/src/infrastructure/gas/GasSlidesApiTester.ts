
export class GasSlidesApiTester {
    /**
     * Executes a prototype batch update to demonstrate API speed and atomicity.
     * @param presentationId The ID of the target presentation
     */
    runPrototype(presentationId: string) {
        const pageId = 'TEST_SLIDE_' + Math.floor(Math.random() * 100000);
        const titleId = pageId + '_TITLE';
        const bodyId = pageId + '_BODY';
        const shapeId = pageId + '_SHAPE';

        // Define the batch operations
        // Note: We use 'any' for Request structure here to avoid complex type matching 
        // with the Advanced Service types if strict checking is enabled, 
        // but ideally this adheres to GoogleAppsScript.Slides.Schema.Request
        const requests: any[] = [
            // 1. Create a generic Slide
            {
                createSlide: {
                    objectId: pageId,
                    slideLayoutReference: {
                        predefinedLayout: 'TITLE_AND_BODY'
                    },
                    placeholderIdMappings: [
                        {
                            layoutPlaceholder: { type: 'TITLE', index: 0 },
                            objectId: titleId
                        },
                        {
                            layoutPlaceholder: { type: 'BODY', index: 0 },
                            objectId: bodyId
                        }
                    ]
                }
            },
            // 2. Populate Title
            {
                insertText: {
                    objectId: titleId,
                    text: '🚀 High-Speed API Generation'
                }
            },
            // 3. Populate Body with multiple bullets
            {
                insertText: {
                    objectId: bodyId,
                    text: 'This entire slide was created in ONE round-trip request.\n\nBenefits:\n- No timeout risks for large decks\n- Atomic (all or nothing)\n- Precise ID control'
                }
            },
            // 4. Create a custom shape (Star)
            {
                createShape: {
                    objectId: shapeId,
                    shapeType: 'STAR_5_POINT',
                    elementProperties: {
                        pageObjectId: pageId,
                        size: {
                            width: { magnitude: 120, unit: 'PT' },
                            height: { magnitude: 120, unit: 'PT' }
                        },
                        transform: {
                            scaleX: 1,
                            scaleY: 1,
                            translateX: 500,
                            translateY: 100,
                            unit: 'PT'
                        }
                    }
                }
            },
            // 5. Add text to shape and style it
            {
                insertText: {
                    objectId: shapeId,
                    text: 'FAST'
                }
            },
            {
                updateTextStyle: {
                    objectId: shapeId,
                    style: {
                        bold: true,
                        fontSize: { magnitude: 18, unit: 'PT' },
                        foregroundColor: {
                            opaqueColor: { themeColor: 'ACCENT1' }
                        }
                    },
                    fields: 'bold,fontSize,foregroundColor'
                }
            }
        ];

        Logger.log('Starting Batch Update...');
        const startTime = new Date().getTime();

        try {
            // @ts-ignore: Slides Advanced Service is available in GAS environment
            const resource = { requests: requests };
            // @ts-ignore
            Slides.Presentations.batchUpdate(resource, presentationId);

            const duration = new Date().getTime() - startTime;
            Logger.log(`✅ Batch Update Complete in ${duration}ms!`);
            return `Generated 1 slide with ${requests.length} operations in ${duration}ms.`;

        } catch (e: any) {
            Logger.log('❌ Error in Batch Update: ' + e.toString());
            throw new Error('API Prototype Failed: ' + e.message);
        }
    }
}
