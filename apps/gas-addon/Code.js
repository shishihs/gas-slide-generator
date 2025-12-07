/**
 * ========================================
 * Google Document to Slide Generator
 * ========================================
 * 
 * このスクリプトは以下の機能を提供します：
 * 1. Convert to JSON: ドキュメントの内容をVertex AIを使ってJSON形式に変換
 * 2. Generate Slide: JSONを基にGoogle Slideを自動生成
 */

// ========================================
// Configuration
// ========================================

/**
 * Get the GCP Project ID from Script Properties
 */
function getGcpProjectId() {
    return PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID');
}

/**
 * Get the Vertex AI location from Script Properties
 */
function getVertexAiLocation() {
    return PropertiesService.getScriptProperties().getProperty('VERTEX_AI_LOCATION') || 'asia-northeast1';
}

/**
 * Get the Vertex AI model from Script Properties
 */
function getVertexAiModel() {
    return PropertiesService.getScriptProperties().getProperty('VERTEX_AI_MODEL') || 'gemini-1.5-flash';
}

/**
 * Get the template Slide ID from Script Properties
 */
function getTemplateSlideId() {
    return PropertiesService.getScriptProperties().getProperty('TEMPLATE_SLIDE_ID');
}

/**
 * Setup Script Properties
 * Run this function once manually to set up the properties.
 */
function setupScriptProperties() {
    const props = {
        'GCP_PROJECT_ID': 'YOUR_GCP_PROJECT_ID',
        'VERTEX_AI_LOCATION': 'asia-northeast1',
        'VERTEX_AI_MODEL': 'gemini-1.5-flash',
        'TEMPLATE_SLIDE_ID': 'YOUR_TEMPLATE_SLIDE_ID',
    };
    PropertiesService.getScriptProperties().setProperties(props);
    Logger.log('Script Properties setup complete.');
}


// ========================================
// Menu & UI
// ========================================

/**
 * onOpen trigger - Adds custom menu when document is opened
 * 
 * パターン: Google Workspace アドオンのメニュー追加パターン
 * - ドキュメントが開かれたタイミングでこの関数が自動実行される
 * - createMenu() でカスタムメニューを作成
 * - addItem() で各メニュー項目と対応する関数を紐付ける
 */
function onOpen() {
    DocumentApp.getUi()
        .createMenu('Slide Generator')
        .addItem('Convert to JSON', 'convertDocumentToJson')
        .addItem('Generate Slide', 'showGenerateSlideSidebar')
        .addSeparator()
        .addItem('Settings', 'showSettingsSidebar')
        .addToUi();
}

/**
 * Show the Generate Slide sidebar
 * 
 * パターン: サイドバー表示のパターン
 * - HtmlService.createHtmlOutputFromFile() でHTMLファイルを読み込む
 * - setTitle() でサイドバーのタイトルを設定
 * - showSidebar() でサイドバーを表示
 */
function showGenerateSlideSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('GenerateSlideSidebar')
        .setTitle('Generate Slide');
    DocumentApp.getUi().showSidebar(html);
}

/**
 * Show the Settings sidebar
 */
function showSettingsSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('SettingsSidebar')
        .setTitle('Settings');
    DocumentApp.getUi().showSidebar(html);
}


// ========================================
// Core Functions
// ========================================

/**
 * Convert the current document to JSON using Vertex AI
 * 
 * This function:
 * 1. Gets the document content
 * 2. Calls Vertex AI directly
 * 3. Receives JSON response
 * 4. Stores the JSON for slide generation
 */
function convertDocumentToJson() {
    const ui = DocumentApp.getUi();

    try {
        // 1. Get document content
        const doc = DocumentApp.getActiveDocument();
        const body = doc.getBody();
        const documentText = body.getText();

        if (!documentText.trim()) {
            ui.alert('Error', 'Document is empty. Please add content before converting.', ui.ButtonSet.OK);
            return null;
        }

        // 2. Call Vertex AI directly
        const jsonResult = callVertexAI(documentText);

        // 3. Store the result
        storeJsonData(jsonResult);

        Logger.log('JSON Result: ' + JSON.stringify(jsonResult));

        ui.alert('Success', 'Document converted to JSON successfully!\n\nClick "Generate Slide" to create slides.', ui.ButtonSet.OK);

        return jsonResult;

    } catch (error) {
        Logger.log('Error: ' + error.toString());
        ui.alert('Error', 'Failed to convert document: ' + error.message, ui.ButtonSet.OK);
        throw error;
    }
}

/**
 * Generate slides from JSON data
 * Called from the sidebar
 * 
 * @param {Object} jsonData - The JSON data containing slide information
 * @returns {Object} - Result containing the new slide URL
 */
function generateSlidesFromJson(jsonData) {
    try {
        const templateId = getTemplateSlideId();

        if (!templateId || templateId === 'YOUR_TEMPLATE_SLIDE_ID') {
            throw new Error('Template Slide ID not configured. Please run Settings first.');
        }

        // 1. Copy the template slide
        const templateFile = DriveApp.getFileById(templateId);
        const newFile = templateFile.makeCopy('Generated Slide - ' + new Date().toISOString());
        const newSlideId = newFile.getId();

        // 2. Open the new presentation
        const presentation = SlidesApp.openById(newSlideId);

        // 3. Process JSON and create slides
        // TODO: Implement slide generation logic based on JSON structure
        // Example structure expected:
        // {
        //   "title": "Presentation Title",
        //   "slides": [
        //     { "type": "title", "title": "...", "subtitle": "..." },
        //     { "type": "content", "title": "...", "bullets": ["...", "..."] },
        //     ...
        //   ]
        // }

        if (jsonData && jsonData.slides) {
            // Process each slide in the JSON
            jsonData.slides.forEach((slideData, index) => {
                // TODO: Add slide content based on slideData
                Logger.log('Processing slide ' + (index + 1) + ': ' + JSON.stringify(slideData));
            });
        }

        // 4. Return the new slide URL
        return {
            success: true,
            slideId: newSlideId,
            slideUrl: 'https://docs.google.com/presentation/d/' + newSlideId + '/edit'
        };

    } catch (error) {
        Logger.log('Error generating slides: ' + error.toString());
        return {
            success: false,
            error: error.message
        };
    }
}


// ========================================
// Vertex AI Direct API Call
// ========================================

/**
 * Call Vertex AI directly from GAS
 * 
 * GAS から Vertex AI を直接呼び出す
 * - ScriptApp.getOAuthToken() で OAuth トークンを取得
 * - UrlFetchApp で Vertex AI REST API を呼び出し
 * - Cloud Functions などのプロキシは不要
 * 
 * @param {string} documentContent - The document content to process
 * @returns {Object} - The JSON response from Vertex AI
 */
function callVertexAI(documentContent) {
    const projectId = getGcpProjectId();
    const location = getVertexAiLocation();
    const model = getVertexAiModel();

    if (!projectId || projectId === 'YOUR_GCP_PROJECT_ID') {
        throw new Error('GCP Project ID not configured. Please run Settings first.');
    }

    // Get OAuth token for GCP API access
    const token = ScriptApp.getOAuthToken();

    // Vertex AI REST API endpoint
    const endpoint = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${model}:generateContent`;

    // Build the prompt
    const prompt = buildSlideGenerationPrompt(documentContent);

    const requestBody = {
        contents: [{
            role: 'user',
            parts: [{
                text: prompt
            }]
        }],
        generationConfig: {
            temperature: 0.2,
            topK: 40,
            topP: 0.95,
            maxOutputTokens: 8192,
        },
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            'Authorization': `Bearer ${token}`
        },
        payload: JSON.stringify(requestBody),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(endpoint, options);
        const code = response.getResponseCode();
        const content = response.getContentText();

        Logger.log(`Response Code: ${code}`);

        if (code !== 200) {
            Logger.log(`Error Response: ${content}`);
            throw new Error(`Vertex AI API returned status ${code}`);
        }

        const result = JSON.parse(content);

        // Extract the generated text
        const generatedText = result.candidates?.[0]?.content?.parts?.[0]?.text;

        if (!generatedText) {
            throw new Error('No content generated from Vertex AI');
        }

        // Parse the JSON from the response
        return parseJsonFromResponse(generatedText);

    } catch (error) {
        Logger.log(`Error: ${error.toString()}`);
        throw error;
    }
}

/**
 * Build the prompt for slide generation
 */
function buildSlideGenerationPrompt(documentContent) {
    return `
You are a presentation designer. Convert the following document content into a structured JSON format for Google Slides.

Document Content:
---
${documentContent}
---

Output the result as a valid JSON object with the following structure:
{
  "title": "Main presentation title",
  "slides": [
    {
      "type": "title",
      "title": "Slide title",
      "subtitle": "Optional subtitle"
    },
    {
      "type": "content",
      "title": "Slide title",
      "bullets": ["Point 1", "Point 2", "Point 3"]
    },
    {
      "type": "section",
      "title": "Section title"
    },
    {
      "type": "conclusion",
      "title": "Conclusion",
      "keyPoints": ["Key takeaway 1", "Key takeaway 2"]
    }
  ]
}

Guidelines:
- Create 5-10 slides based on the content
- Use clear, concise bullet points
- Include a title slide at the beginning
- Group related content into sections
- End with a conclusion or summary slide

Output ONLY the JSON, no additional text or explanation.
`;
}

/**
 * Parse JSON from Vertex AI response
 */
function parseJsonFromResponse(text) {
    let jsonString = text.trim();

    // Remove markdown code blocks if present
    if (jsonString.startsWith('```json')) {
        jsonString = jsonString.slice(7);
    } else if (jsonString.startsWith('```')) {
        jsonString = jsonString.slice(3);
    }

    if (jsonString.endsWith('```')) {
        jsonString = jsonString.slice(0, -3);
    }

    jsonString = jsonString.trim();

    try {
        return JSON.parse(jsonString);
    } catch (error) {
        Logger.log('Failed to parse JSON: ' + error.toString());
        Logger.log('Raw text: ' + text);
        throw new Error('Failed to parse JSON from Vertex AI response');
    }
}


// ========================================
// Utility Functions
// ========================================

/**
 * Get the current document's content as plain text
 * Called from sidebar
 */
function getDocumentContent() {
    const doc = DocumentApp.getActiveDocument();
    return doc.getBody().getText();
}

/**
 * Store JSON data in document properties
 * @param {Object} jsonData - The JSON data to store
 */
function storeJsonData(jsonData) {
    const jsonString = JSON.stringify(jsonData);
    PropertiesService.getDocumentProperties().setProperty('SLIDE_JSON', jsonString);
}

/**
 * Retrieve stored JSON data from document properties
 * @returns {Object|null} - The stored JSON data or null if not found
 */
function getStoredJsonData() {
    const jsonString = PropertiesService.getDocumentProperties().getProperty('SLIDE_JSON');
    return jsonString ? JSON.parse(jsonString) : null;
}

/**
 * Save script properties from the Settings sidebar
 * @param {Object} props - Object containing property key-value pairs
 */
function saveScriptProperties(props) {
    PropertiesService.getScriptProperties().setProperties(props);
    Logger.log('Script Properties saved: ' + Object.keys(props).join(', '));
}

/**
 * Get current script properties for display in Settings sidebar
 * @returns {Object} - Current property values
 */
function getCurrentSettings() {
    const props = PropertiesService.getScriptProperties();
    return {
        gcpProjectId: props.getProperty('GCP_PROJECT_ID') || '',
        vertexAiLocation: props.getProperty('VERTEX_AI_LOCATION') || 'asia-northeast1',
        vertexAiModel: props.getProperty('VERTEX_AI_MODEL') || 'gemini-1.5-flash',
        templateSlideId: props.getProperty('TEMPLATE_SLIDE_ID') || ''
    };
}
