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
    return PropertiesService.getScriptProperties().getProperty('VERTEX_AI_MODEL') || 'gemini-2.5-flash';
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
        'VERTEX_AI_MODEL': 'gemini-2.5-flash',
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
 * 4. Stores the result and opens the sidebar
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

        doc.saveAndClose(); // Save pending changes

        // 2. Call Vertex AI directly
        const jsonResult = callVertexAI(documentText);

        // 3. Store the result
        storeJsonData(jsonResult);

        Logger.log('JSON Result: ' + JSON.stringify(jsonResult));

        // 4. Open the sidebar automatically to show the result
        showGenerateSlideSidebar();

        return jsonResult;

    } catch (error) {
        Logger.log('Error: ' + error.toString());
        ui.alert('Error', 'Failed to convert document: ' + error.message, ui.ButtonSet.OK);
        throw error;
    }
}

/**
 * Get the Slide Generator API URL from Script Properties
 */
function getSlideGeneratorApiUrl() {
    return PropertiesService.getScriptProperties().getProperty('SLIDE_GENERATOR_API_URL');
}

/**
 * Generate slides from JSON data using the external Slide Generator API
 * Called from the sidebar
 * 
 * @param {Object} jsonData - The JSON data containing slide information
 * @returns {Object} - Result containing the new slide URL
 */
function generateSlidesFromJson(jsonData) {
    try {
        const apiUrl = getSlideGeneratorApiUrl();

        if (!apiUrl || apiUrl === 'YOUR_SLIDE_GENERATOR_API_URL') {
            throw new Error('Slide Generator API URL not configured. Please run Settings first.');
        }

        // Get ID Token for authentication
        const token = ScriptApp.getIdentityToken();
        if (!token) {
            // If checking fails, might need to adjust manifest oauth scopes (openid)
            Logger.log('Warning: No ID Token available. Trying anonymous access or ensure openid scope.');
        }

        const templateId = getTemplateSlideId();

        // Flatten the payload so api.ts receives everything in one object
        const payload = {
            ...jsonData,
            templateId: templateId || undefined
        };

        const headers = token ? { 'Authorization': `Bearer ${token}` } : {};

        const options = {
            method: 'post',
            contentType: 'application/json',
            headers: headers,
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(apiUrl, options);
        const code = response.getResponseCode();
        const content = response.getContentText();

        if (code !== 200) {
            throw new Error(`Slide Generator API returned ${code}: ${content}`);
        }

        const result = JSON.parse(content);

        // Expected result format: { "url": "https://docs.google.com/presentation/..." }
        if (!result.url) {
            throw new Error('Invalid response from API: URL not found');
        }

        return {
            success: true,
            slideUrl: result.url
        };

    } catch (error) { // Type assertion
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
    // === MOCK MODE FOR TESTING ===
    const userProperties = PropertiesService.getUserProperties();
    const useMock = userProperties.getProperty('USE_MOCK_RESPONSE') === 'true';

    if (useMock) {
        Logger.log('Using Mock Vertex AI Response');
        return {
            "title": "Mock Presentation Title",
            "slides": [
                {
                    "layout": "TITLE",
                    "title": "Mock Presentation",
                    "subtitle": "Generated by Mock Mode",
                    "content": [],
                    "notes": "Welcome"
                },
                {
                    "layout": "CONTENT",
                    "title": "Mock Slide 1",
                    "content": ["Mock Point 1", "Mock Point 2"],
                    "notes": "Generated by Mock Mode"
                },
                {
                    "layout": "TITLE_AND_BODY",
                    "title": "Mock Slide 2",
                    "content": ["Another point"],
                    "notes": "Notes here"
                }
            ],
            "theme": "DEFAULT"
        };
    }
    // =============================

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
            temperature: 0.1, // Lower temperature for more deterministic output (better for JSON)
            topK: 40,
            topP: 0.95,
            maxOutputTokens: 8192,
            responseMimeType: "application/json" // Force JSON output mode if supported by the model
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
You are a professional presentation designer. Your task is to convert the following document content into a structured JSON format for Google Slides.

Target Audience: Business professionals
Tone: Professional, Clear, Concise

Document Content:
---
${documentContent}
---

Create a presentation structure with 5-12 slides depending on the content length.
The output MUST be a valid JSON object following this strict schema:

{
  "theme": "modern", // Suggested theme keyword (e.g., modern, corporate, creative)
  "metada": {
    "estimatedDuration": "10 min",
    "audience": "General"
  },
  "slides": [
    {
      "type": "TITLE",
      "title": "Main Presentation Title",
      "subtitle": "Subtitle or Presenter Name",
      "notes": "Speaker notes for the title slide"
    },
    {
      "type": "AGENDA",
      "title": "Agenda",
      "items": ["Topic 1", "Topic 2", "Topic 3"]
    },
    {
      "type": "CONTENT",
      "title": "Slide Title",
      "body": {
        "text": "Main paragraph text if any",
        "bullets": ["Key point 1", "Key point 2", "Key point 3"]
      },
      "notes": "Speaker notes explaining the content"
    },
    {
      "type": "SECTION",
      "title": "Section Title",
      "subtitle": "Brief description of the section"
    },
    {
      "type": "CONCLUSION",
      "title": "Summary / Key Takeaways",
      "points": ["Takeaway 1", "Takeaway 2"]
    }
  ]
}

Rules:
1. **Summarize**: Do not just copy text. Summarize long paragraphs into concise bullet points.
2. **Structure**: logical flow (Title -> Agenda -> Content -> Conclusion).
3. **Content**: Each slide should have 3-5 bullet points max for readability.
4. **JSON Only**: Output pure JSON. No markdown formatting (like \`\`\`json), no introductory text.
`;
}

/**
 * Parse JSON from Vertex AI response
 * More robust parsing to handle accumulated markdown or preamble
 */
function parseJsonFromResponse(text) {
    let jsonString = text.trim();

    // 1. Remove markdown code blocks if present
    // Matches ```json ... ``` or just ``` ... ```
    const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)\s*```/i;
    const match = jsonString.match(codeBlockRegex);
    if (match && match[1]) {
        jsonString = match[1].trim();
    }

    // 2. If it still looks like it has extra text, try to find the first { and last }
    const firstBrace = jsonString.indexOf('{');
    const lastBrace = jsonString.lastIndexOf('}');

    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
        jsonString = jsonString.substring(firstBrace, lastBrace + 1);
    }

    try {
        return JSON.parse(jsonString);
    } catch (error) {
        Logger.log('Failed to parse JSON: ' + error.toString());
        Logger.log('Raw text: ' + text);
        Logger.log('Attempted string: ' + jsonString);
        throw new Error('Failed to parse JSON from Vertex AI response. The AI might have returned invalid format.');
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
    const scriptProps = { ...props };
    delete scriptProps.USE_MOCK_RESPONSE;

    PropertiesService.getScriptProperties().setProperties(scriptProps);

    if (props.USE_MOCK_RESPONSE !== undefined) {
        PropertiesService.getUserProperties().setProperty('USE_MOCK_RESPONSE', props.USE_MOCK_RESPONSE);
    }

    Logger.log('Properties saved.');
}

/**
 * Get current script properties for display in Settings sidebar
 * @returns {Object} - Current property values
 */
function getCurrentSettings() {
    const props = PropertiesService.getScriptProperties();
    const userProps = PropertiesService.getUserProperties();
    return {
        gcpProjectId: props.getProperty('GCP_PROJECT_ID') || '',
        vertexAiLocation: props.getProperty('VERTEX_AI_LOCATION') || 'asia-northeast1',
        vertexAiModel: props.getProperty('VERTEX_AI_MODEL') || 'gemini-1.5-flash',
        templateSlideId: props.getProperty('TEMPLATE_SLIDE_ID') || '',
        slideGeneratorApiUrl: props.getProperty('SLIDE_GENERATOR_API_URL') || '',
        useMockResponse: userProps.getProperty('USE_MOCK_RESPONSE') || 'false'
    };
}
