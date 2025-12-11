
import { test, expect } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import { execSync } from 'child_process';

const previewPath = path.resolve(__dirname, '../preview.html');

test.describe('Visual Regression', () => {
    test.beforeAll(() => {
        // Ensure preview exists
        if (!fs.existsSync(previewPath)) {
            console.log('Generating preview...');
            execSync('npm run preview', { cwd: path.resolve(__dirname, '..') });
        }
    });

    test('capture slides', async ({ page }) => {
        const fileUrl = `file://${previewPath}`;
        await page.goto(fileUrl);

        // Wait for content (SVGs)
        await page.waitForSelector('.slide-container');

        const slides = await page.locator('.slide-container').all();
        console.log(`Found ${slides.length} slides.`);

        for (let i = 0; i < slides.length; i++) {
            const slide = slides[i];
            const label = await slide.locator('.slide-label').innerText();
            const slideEl = slide.locator('.slide');

            // Clean filename
            const name = label.replace(/[^a-z0-9]/gi, '_').toLowerCase();

            await slideEl.screenshot({ path: `tests/visual-results/${name}.png` });
        }
    });
});
