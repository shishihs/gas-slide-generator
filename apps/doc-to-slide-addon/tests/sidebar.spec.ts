
import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';

test('Generate Slide Sidebar UI Flow', async ({ page }) => {
    // 1. Read the actual Sidebar HTML
    const sidebarHtmlPath = path.join(__dirname, '../src/GenerateSlideSidebar.html');
    const sidebarHtmlContent = fs.readFileSync(sidebarHtmlPath, 'utf8');

    // 2. Read the Mock Wrapper
    const mockWrapperPath = path.join(__dirname, 'mock-sidebar.html');
    const mockWrapperContent = fs.readFileSync(mockWrapperPath, 'utf8');

    // 3. Combine them: inject sidebar content into the body of the mock wrapper
    // We'll just append it or replace the container. 
    // Since the Sidebar HTML is a full <html> document, we need to extract body or just dump it.
    // A simple way is to use `page.setContent` with the Mock script appended to the real Sidebar.

    // STRATEGY CHANGE: 
    // Instead of injecting into `mock-sidebar.html`, let's take the Real Sidebar HTML 
    // and inject the Mock Script at the top of <head> or right before </body>.
    // This ensures we test the Real HTML exactly.

    // Mock script to define google.script.run
    const mockScript = `
    <script>
    window.google = {
      script: {
        run: {
          withSuccessHandler: function(cb) { this._success = cb; return this; },
          withFailureHandler: function(cb) { this._failure = cb; return this; },
          convertDocumentToJson: function() {
            setTimeout(() => {
               if(this._success) this._success({
                 title: "Mock Pres",
                 slides: [{type:"TITLE", title:"Test"}]
               });
            }, 100);
          },
          getStoredJsonData: function() {
             setTimeout(() => {
               if(this._success) this._success(null);
             }, 50);
          },
          generateSlidesFromJson: function(json) {
             setTimeout(() => {
               if(this._success) this._success({
                 success: true,
                 slideUrl: "https://example.com/presentation"
               });
             }, 500);
          }
        }
      }
    };
    </script>
  `;

    // Insert mock script before the first script tag or in head
    const finalHtml = sidebarHtmlContent.replace('<head>', '<head>' + mockScript);

    // 4. Load the page
    await page.setContent(finalHtml);

    // 5. Interact with UI

    // Initial state: "Step 1" should be visible
    await expect(page.locator('text=Step 1: Prepare JSON Data')).toBeVisible();

    // "Step 2" and "Step 3" should be hidden initially (if no stored json)
    // (The mock getStoredJsonData returns null)
    await expect(page.locator('#previewSection')).toBeHidden();

    // Click "Convert Document to JSON"
    await page.click('button:has-text("Convert Document to JSON")');

    // Should see loading status? (Optional check)

    // Wait for "Step 2" to appear (onConvertSuccess)
    await expect(page.locator('#previewSection')).toBeVisible();
    await expect(page.locator('#jsonPreview')).toContainText('Mock Pres');

    // "Step 3" should also be visible now
    await expect(page.locator('#generateSection')).toBeVisible();

    // Click "Generate Slides"
    await page.click('#generateBtn');

    // Button should be disabled during generation
    await expect(page.locator('#generateBtn')).toBeDisabled();
    await expect(page.locator('#generateBtn')).toHaveText('Generating...');

    // Wait for success
    await expect(page.locator('#resultSection')).toBeVisible();
    await expect(page.locator('#slideLink')).toHaveAttribute('href', 'https://example.com/presentation');

    // Button re-enabled
    await expect(page.locator('#generateBtn')).toBeEnabled();
    await expect(page.locator('#generateBtn')).toHaveText('Generate Slides');
});
