const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

(async () => {
    console.log('Launching browser...');
    const browser = await puppeteer.launch({
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();

    // Set viewport to HD (1280x720) with high DPI for better quality text
    await page.setViewport({ width: 1280, height: 720, deviceScaleFactor: 2 });

    const filePath = path.join(__dirname, 'preview.html');
    const fileUrl = `file://${filePath}`;
    console.log(`Opening ${fileUrl}...`);

    try {
        await page.goto(fileUrl, { waitUntil: 'networkidle0' });
    } catch (e) {
        console.error('Error opening page:', e);
        process.exit(1);
    }

    // Wait for at least one slide to be present
    try {
        await page.waitForSelector('.slide-container', { timeout: 5000 });
    } catch (e) {
        console.error('Timeout waiting for .slide-container. Page content might be empty or render.js failed.');
        // Take a screenshot of the error state
        await page.screenshot({ path: 'error_state.png' });
        process.exit(1);
    }

    const slides = await page.$$('.slide-container');
    console.log(`Found ${slides.length} slides.`);

    const outputDir = path.join(__dirname, 'slide_sample');
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    for (let i = 0; i < slides.length; i++) {
        const slide = slides[i];
        const fileName = `slide_${(i + 1).toString().padStart(2, '0')}.png`;
        const outputPath = path.join(outputDir, fileName);

        // Screenshot only the slide element
        await slide.screenshot({ path: outputPath });
        console.log(`Saved ${fileName}`);
    }

    await browser.close();
    console.log('Done.');
})();
