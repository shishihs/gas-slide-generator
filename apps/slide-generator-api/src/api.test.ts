
import { describe, it, expect, vi } from 'vitest';
import { generateSlides } from './api';
import { GasSlideRepository } from './infrastructure/gas/GasSlideRepository';
import { PresentationApplicationService } from './application/PresentationApplicationService';

// Mock the dependencies
vi.mock('./infrastructure/gas/GasSlideRepository');
vi.mock('./application/PresentationApplicationService');

// Mock GAS Logger
global.Logger = {
    log: (msg: string) => console.log('[GAS LOG]:', msg),
} as any;

describe('generateSlides', () => {
    it('should handle root-level theme (New Format)', () => {
        const mockCreatePresentation = vi.fn().mockReturnValue('mock-url');
        // @ts-ignore
        PresentationApplicationService.mockImplementation(function () {
            return {
                createPresentation: mockCreatePresentation
            };
        });

        const input = {
            theme: 'Blue',
            slides: [
                { type: 'title', title: 'Test' }
            ]
        };

        const result = generateSlides(input);

        expect(result.success).toBe(true);
        expect(result.url).toBe('mock-url');

        // Check if settings.theme was passed
        const callArgs = mockCreatePresentation.mock.calls[0][0];
        expect(callArgs.settings.theme).toBe('Blue');
    });

    it('should handle array input (Legacy Format)', () => {
        const mockCreatePresentation = vi.fn().mockReturnValue('mock-url');
        // @ts-ignore
        PresentationApplicationService.mockImplementation(function () {
            return {
                createPresentation: mockCreatePresentation
            };
        });

        const input = [
            { type: 'title', title: 'Test' }
        ];

        const result = generateSlides(input);

        expect(result.success).toBe(true);
        // Validates backward compatibility (no crash)
        const callArgs = mockCreatePresentation.mock.calls[0][0];
        expect(callArgs.slides).toHaveLength(1);
    });

    it('should check settings.theme if root theme is missing', () => {
        const mockCreatePresentation = vi.fn().mockReturnValue('mock-url');
        // @ts-ignore
        PresentationApplicationService.mockImplementation(function () {
            return {
                createPresentation: mockCreatePresentation
            };
        });

        const input = {
            settings: { theme: 'Red' },
            slides: [{ type: 'title', title: 'Test' }]
        };

        const result = generateSlides(input);

        const callArgs = mockCreatePresentation.mock.calls[0][0];
        expect(callArgs.settings.theme).toBe('Red');
    });

    it('should prioritize root theme over settings.theme', () => {
        const mockCreatePresentation = vi.fn().mockReturnValue('mock-url');
        // @ts-ignore
        PresentationApplicationService.mockImplementation(function () {
            return {
                createPresentation: mockCreatePresentation
            };
        });

        const input = {
            theme: 'Green',
            settings: { theme: 'Red' }, // Should be ignored
            slides: [{ type: 'title', title: 'Test' }]
        };

        const result = generateSlides(input);

        const callArgs = mockCreatePresentation.mock.calls[0][0];
        expect(callArgs.settings.theme).toBe('Green');
    });
});
