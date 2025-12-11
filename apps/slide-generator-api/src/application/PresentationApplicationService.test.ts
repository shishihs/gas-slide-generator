import { vi, describe, it, expect, beforeEach, Mock } from 'vitest';
import { PresentationApplicationService } from './PresentationApplicationService';
import { ISlideRepository } from '../domain/repositories/ISlideRepository';
import { Presentation } from '../domain/model/Presentation';

describe('PresentationApplicationService', () => {
    let service: PresentationApplicationService;
    let mockRepo: any;

    beforeEach(() => {
        mockRepo = {
            createPresentation: vi.fn().mockReturnValue('https://mock-url.com')
        };
        service = new PresentationApplicationService(mockRepo);
    });

    it('should orchestrate presentation creation', () => {
        const request = {
            title: 'My Presentation',
            templateId: 'tmpl-123',
            slides: [
                { title: 'Slide 1', content: ['A', 'B'], layout: 'TITLE' },
                { title: 'Slide 2', content: ['C'], notes: 'Speaker notes' }
            ]
        };

        const url = service.createPresentation(request);

        expect(url).toBe('https://mock-url.com');
        expect(mockRepo.createPresentation).toHaveBeenCalledTimes(1);
        expect(mockRepo.createPresentation).toHaveBeenCalledWith(expect.any(Presentation), 'tmpl-123', undefined, undefined);

        // Verify Domain Object construction
        const capturedPresentation = mockRepo.createPresentation.mock.calls[0][0];
        expect(capturedPresentation).toBeInstanceOf(Presentation);
        expect(capturedPresentation.title).toBe('My Presentation');
        expect(capturedPresentation.slides).toHaveLength(2);
        expect(capturedPresentation.slides[0].layout).toBe('TITLE');
        expect(capturedPresentation.slides[1].layout).toBe('CONTENT'); // Default
        expect(capturedPresentation.slides[1].notes).toBe('Speaker notes');
    });
});
