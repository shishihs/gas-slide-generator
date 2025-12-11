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
    const val = PropertiesService.getScriptProperties().getProperty('TEMPLATE_SLIDE_ID');
    return extractPresentationIdFromUrl(val);
}

/**
 * Get the Gem URL for manual JSON generation advice
 */
function getGemUrl() {
    return PropertiesService.getScriptProperties().getProperty('GEM_URL') || '';
}

/**
 * Check if the provided slide ID is accessible
 * @param {string} id - The Slide ID or URL
 */
function checkSlideIdAccess(id) {
    const cleanId = extractPresentationIdFromUrl(id);
    if (!cleanId) {
        return { valid: false, message: 'IDが空、または無効です' };
    }

    try {
        const pres = SlidesApp.openById(cleanId);
        const name = pres.getName();
        return { valid: true, message: `✅ アクセス成功: ${name}` };
    } catch (e) {
        return { valid: false, message: `❌ アクセス失敗: ${e.message}` };
    }
}

/**
 * Extract Presentation ID from a URL or return as is
 * @param {string} urlOrId - The URL or ID string
 * @returns {string|null} - The extracted ID or original string
 */
function extractPresentationIdFromUrl(urlOrId) {
    if (!urlOrId) return null;
    // Common pattern for Google Doc/Slide/Sheet URLs
    // e.g. https://docs.google.com/presentation/d/1234567890abcdef/edit...
    const match = urlOrId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : urlOrId;
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
        'GEM_URL': 'https://gemini.google.com/app/...'
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
        .createMenu('スライド起草くん')
        .addItem('スライド生成', 'showGenerateSlideSidebar')
        .addItem('デザイン設定', 'showDesignSettingsSidebar')
        .addItem('設定', 'showSettingsSidebar')
        .addSeparator()
        .addItem('ヘルプ', 'showHelpSidebar')
        .addItem('デバッグ情報', 'debugTest')
        .addToUi();
}

/**
 * Diagnostic function to debug permissions and file access
 */
function debugTest() {
    const ui = DocumentApp.getUi();
    const logs = [];
    const log = (msg) => {
        Logger.log(msg);
        logs.push(msg);
    };

    try {
        try {
            log('User: ' + Session.getActiveUser().getEmail());
            log('EffectiveUser: ' + Session.getEffectiveUser().getEmail());
        } catch (e) {
            log('User Info Error: ' + e.toString());
        }

        // Check DriveApp generic access
        try {
            const root = DriveApp.getRootFolder();
            log('Drive Root Access: OK (' + root.getName() + ')');
        } catch (e) {
            log('Drive Root Error: ' + e.toString());
        }

        // Check Template
        const templateId = getTemplateSlideId();
        log('Template ID: ' + templateId);

        if (templateId) {
            try {
                log('Attempting verify copy (Advanced Drive API)...');
                // Use Advanced Drive Service for debug as well
                const copy = Drive.Files.copy({
                    title: 'Debug_Copy_Test',
                    parents: [{ id: 'root' }]
                }, templateId);

                log('Copy Success: ' + copy.id);

                // Cleanup using Drive.Files since DriveApp might be broken
                try {
                    Drive.Files.trash(copy.id);
                    log('Cleanup: OK');
                } catch (cleanupErr) {
                    log('Cleanup Warning: ' + cleanupErr.toString());
                }

            } catch (e) {
                log('Template Copy Error (Advanced API): ' + e.toString());
            }
        } else {
            log('No Template ID set.');
        }

    } catch (e) {
        log('Fatal Error: ' + e.toString());
    }

    ui.alert('Debug Log', logs.slice(0, 30).join('\n'), ui.ButtonSet.OK);
}

/**
 * Create a template slide presentation with all required layouts.
 * The template ID will be saved to settings automatically.
 * 
 * Uses Google Slides PredefinedLayouts to ensure proper placeholder structure.
 * Each slide is tagged with a layout name in speaker notes for identification.
 * 
 * Required layouts based on LayoutType:
 * - TITLE: Title slide with title and subtitle (uses TITLE layout)
 * - AGENDA: Agenda/outline slide (uses TITLE_AND_BODY layout)
 * - CONTENT: Main content slide with title and body (uses TITLE_AND_BODY layout)
 * - SECTION: Section divider slide (uses SECTION_HEADER layout)
 * - CONCLUSION: Closing/summary slide (uses TITLE_AND_BODY layout)
 */
function createTemplateSlide() {
    const ui = DocumentApp.getUi();

    try {
        const result = createTemplatePresentation();

        // Show success message with instructions
        ui.alert(
            '✅ テンプレート作成完了',
            `テンプレートスライドを作成し、設定に保存しました。\n` +
            `作成予定: ${result.expected} 枚\n` +
            `作成成功: ${result.count} 枚\n\n` +
            `【作成されたレイアウト】\n` +
            `• TITLE - 表紙\n` +
            `• AGENDA - 目次\n` +
            `• CONTENT - メインコンテンツ\n` +
            `• SECTION - セクション区切り\n` +
            `• CONCLUSION - まとめ\n\n` +
            `【次のステップ】\n` +
            `1. 以下のURLでテンプレートを開く\n` +
            `2. スライド > テーマを変更 でデザインを選択\n` +
            `3. フォントや色を調整\n` +
            `4. Add-on で Generate Slide を実行\n\n` +
            `URL: ${result.url}`,
            ui.ButtonSet.OK
        );

        Logger.log('Template created: ' + result.id);
        return result;

    } catch (error) {
        ui.alert('エラー', 'テンプレート作成に失敗しました: ' + error.toString(), ui.ButtonSet.OK);
        Logger.log('Error creating template: ' + error.toString());
        throw error;
    }
}

/**
 * Core logic to create the template presentation
 * Returns information about the created template
 */
function createTemplatePresentation() {
    // 1. Create a new presentation and get ID
    Logger.log('Creating new presentation...');
    let presentation = SlidesApp.create('Slide Generator Template');
    const presentationId = presentation.getId();
    const presentationUrl = presentation.getUrl();

    // 2. Modify the presentation directly without closing/reopening
    // This avoids race conditions where the ID is not yet propagated.

    const defaultSlides = presentation.getSlides();
    // Remove the default blank slide
    for (const slide of defaultSlides) {
        slide.remove();
    }

    // Define layouts with their corresponding PredefinedLayout and structure
    const layoutConfigs = [
        {
            name: 'TITLE',
            predefinedLayout: SlidesApp.PredefinedLayout.TITLE,
            description: 'プレゼンテーションの表紙（タイトル＋サブタイトル）',
            placeholders: { title: 'タイトルを入力', subtitle: 'サブタイトルを入力' }
        },
        {
            name: 'AGENDA',
            predefinedLayout: SlidesApp.PredefinedLayout.TITLE_AND_BODY,
            description: '目次・アジェンダ用（タイトル＋本文）',
            placeholders: { title: 'アジェンダ', body: '• 項目1\n• 項目2\n• 項目3' }
        },
        {
            name: 'CONTENT',
            predefinedLayout: SlidesApp.PredefinedLayout.TITLE_AND_BODY,
            description: 'メインコンテンツ用（タイトル＋本文）',
            placeholders: { title: 'スライドタイトル', body: '• ポイント1\n• ポイント2\n• ポイント3' }
        },
        {
            name: 'SECTION',
            predefinedLayout: SlidesApp.PredefinedLayout.SECTION_HEADER,
            description: 'セクション区切り用（大きなタイトル）',
            placeholders: { title: 'セクション名' }
        },
        {
            name: 'CONCLUSION',
            predefinedLayout: SlidesApp.PredefinedLayout.TITLE_AND_BODY,
            description: '結論・まとめ用（タイトル＋本文）',
            placeholders: { title: 'まとめ', body: '• 結論1\n• 結論2' }
        }
    ];

    // Create slides for each layout
    Logger.log('Creating ' + layoutConfigs.length + ' layout slides...');

    let createdCount = 0;
    for (const config of layoutConfigs) {
        Logger.log('Creating slide for layout: ' + config.name);
        try {
            const slide = presentation.appendSlide(config.predefinedLayout);
            createdCount++;

            // Tag the slide with layout name in speaker notes (critical for identification)
            slide.getNotesPage().getSpeakerNotesShape().getText().setText(
                `LAYOUT:${config.name}\n${config.description}`
            );

            // Fill in placeholder text
            const shapes = slide.getShapes();
            for (const shape of shapes) {
                const placeholderType = shape.getPlaceholderType();

                if (placeholderType === SlidesApp.PlaceholderType.TITLE ||
                    placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
                    if (config.placeholders.title) {
                        shape.getText().setText(config.placeholders.title);
                    }
                } else if (placeholderType === SlidesApp.PlaceholderType.SUBTITLE) {
                    if (config.placeholders.subtitle) {
                        shape.getText().setText(config.placeholders.subtitle);
                    }
                } else if (placeholderType === SlidesApp.PlaceholderType.BODY) {
                    if (config.placeholders.body) {
                        shape.getText().setText(config.placeholders.body);
                    }
                }
            }

            // Add a small label at the bottom to identify the layout (for user reference)
            const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 20, 500, 300, 20);
            labelShape.getText().setText(`[TEMPLATE: ${config.name}]`);
            labelShape.getText().getTextStyle().setFontSize(8).setForegroundColor('#AAAAAA');
        } catch (e) {
            Logger.log('Error creating slide for ' + config.name + ': ' + e.toString());
        }
    }

    // Finalize changes
    presentation.saveAndClose();

    // Save to settings
    PropertiesService.getScriptProperties().setProperty('TEMPLATE_SLIDE_ID', presentationId);

    return {
        id: presentationId,
        url: presentationUrl,
        count: createdCount,
        expected: layoutConfigs.length
    };
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
        .setTitle('スライド生成');
    DocumentApp.getUi().showSidebar(html);
}

/**
 * Show the Settings sidebar
 */
function showSettingsSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('SettingsSidebar')
        .setTitle('設定');
    DocumentApp.getUi().showSidebar(html);
}

/**
 * Show the Design Settings sidebar
 */
function showDesignSettingsSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('DesignSettingsSidebar')
        .setTitle('デザイン設定');
    DocumentApp.getUi().showSidebar(html);
}

/**
 * Show the Help sidebar
 */
function showHelpSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('HelpSidebar')
        .setTitle('ヘルプ & ガイド');
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
            ui.alert('エラー', 'ドキュメントが空です。コンテンツを追加してから変換してください。', ui.ButtonSet.OK);
            return null;
        }

        doc.saveAndClose(); // Save pending changes

        // 2. Try to parse the document as a JSON directly (Manual JSON Mode)
        let directJson = null;
        try {
            // Attempt to parse strictly or loosely
            directJson = parseJsonFromResponse(documentText);
        } catch (ignore) {
            // Not a JSON, proceed to AI generation
        }

        // Validate if the parsed JSON looks like our schema
        if (isValidSlideJson(directJson)) {
            Logger.log('Valid JSON structure detected in document. Skipping Vertex AI.');
            storeJsonData(directJson);
            // ui.alert('Success', 'Document content identified as JSON. Slide data loaded directly.', ui.ButtonSet.OK);
            return directJson;
        }

        // 3. Check for Mock Mode or GCP Project ID
        // If mock mode is enabled, we skip the Project ID check
        const userProperties = PropertiesService.getUserProperties();
        const useMock = userProperties.getProperty('USE_MOCK_RESPONSE') === 'true';
        const projectId = getGcpProjectId();

        if (!useMock && (!projectId || projectId === 'YOUR_GCP_PROJECT_ID')) {
            // Return a special result - sidebar will show the manual input section
            return { __manualModeRequired: true };
        }

        // 4. Call Vertex AI directly (Standard Mode or Mock)
        const jsonResult = callVertexAI(documentText);

        // 5. Check for user color settings
        const colors = getUserColors();
        if (colors) {
            if (!jsonResult.settings) jsonResult.settings = {};
            jsonResult.settings.colors = colors;
        }

        storeJsonData(jsonResult);

        Logger.log('JSON Result: ' + JSON.stringify(jsonResult));

        return jsonResult;

    } catch (error) {
        Logger.log('Error: ' + error.toString());
        // Only show alert if it wasn't already handled (e.g. by our own throw)
        if (error.message.indexOf('GCP Project ID missing') === -1) {
            ui.alert('エラー', 'ドキュメントの変換に失敗しました: ' + error.message, ui.ButtonSet.OK);
        }
        throw error;
    }
}

/**
 * Check via UI if Vertex AI Generation is available (Hidden Mode Check)
 */
function getGenerativeModeStatus() {
    const userProperties = PropertiesService.getUserProperties();
    const useMock = userProperties.getProperty('USE_MOCK_RESPONSE') === 'true';
    const projectId = getGcpProjectId();

    const isConfigured = (projectId && projectId !== 'YOUR_GCP_PROJECT_ID');

    return {
        enabled: useMock || isConfigured,
        useMock: useMock
    };
}

/**
 * Validate if the object has the minimum required structure for slides
 * @param {any} data 
 * @returns {boolean}
 */
function isValidSlideJson(data) {
    if (!data || typeof data !== 'object') return false;

    // Case 1: Standard wrapped structure { slides: [...] }
    if (Array.isArray(data.slides)) {
        // Must have at least one valid-looking slide or be an empty array
        if (data.slides.length === 0) return true;

        // Check first slide for minimal required fields
        const first = data.slides[0];
        if (first && (first.title || first.type || first.layout)) {
            return true;
        }
    }

    // Case 2: Array root [ ... ]
    if (Array.isArray(data)) {
        if (data.length === 0) return true;

        const first = data[0];
        // Check if elements look like slides (have title, type, or layout)
        if (first && typeof first === 'object' && (first.title || first.type || first.layout)) {
            return true;
        }
    }

    return false;
}

/**
 * Generate slides from JSON data using the external Slide Generator API Library
 * Called from the sidebar
 * 
 * @param {Object} jsonData - The JSON data containing slide information
 * @returns {Object} - Result containing the new slide URL
 */
function generateSlidesFromJson(jsonData) {
    try {
        let templateId = getTemplateSlideId();
        let destinationId = undefined;



        let slidesData = [];
        let title = 'Generated Presentation'; // Initialize title here

        if (templateId) {
            // Update title based on jsonData for the copy operation
            if (jsonData && jsonData.title) title = jsonData.title;
            else if (Array.isArray(jsonData) && jsonData.length > 0 && jsonData[0].title) title = jsonData[0].title;
            else if (jsonData && jsonData.slides && jsonData.slides.length > 0 && jsonData.slides[0].title) title = jsonData.slides[0].title;

            try {
                Logger.log('Step 1: Preparing to copy using Advanced Drive API (Drive.Files)...');

                Logger.log('Step 2: Copying template ID: ' + templateId + ' with title: ' + title);

                // Use Advanced Drive Service (v2)
                // This bypasses DriveApp Server Errors
                const copiedFile = Drive.Files.copy({
                    title: title,
                    parents: [{ id: 'root' }] // Explicitly place in root or let it inherit? 'root' is safer for finding it.
                }, templateId);

                destinationId = copiedFile.id;
                Logger.log('Step 2 Success: Copy made. Destination ID: ' + destinationId);

            } catch (e) {
                Logger.log('Advanced Drive API Failed: ' + e.toString());
                Logger.log('Falling back to DriveApp (just in case)...');
                try {
                    const templateFile = DriveApp.getFileById(templateId);
                    const newFile = templateFile.makeCopy(title);
                    destinationId = newFile.getId();
                    Logger.log('Fallback Success: Destination ID: ' + destinationId);
                } catch (e2) {
                    Logger.log('All copy methods failed. Error: ' + e2.toString());
                }
            }

            if (!destinationId) {
                Logger.log('Template access failed. Throwing error...');
                throw new Error(
                    `Failed to access or copy the template slide (ID: ${templateId}). \n` +
                    `Please check the ID in settings and ensure you have permission to access it.`
                );
            }
        }


        let theme = 'SIMPLE_LIGHT';

        if (Array.isArray(jsonData)) {
            // If Mock returns an array directly, treat it as the slides list
            slidesData = jsonData;
            // Try to set title from the first slide if available
            if (slidesData.length > 0 && slidesData[0].title) {
                title = slidesData[0].title;
            }
        } else if (jsonData && jsonData.slides && Array.isArray(jsonData.slides)) {
            // Standard format
            slidesData = jsonData.slides;
            title = jsonData.title || title;
            theme = jsonData.theme || theme;
        } else {
            // Fallback/Invalid
            Logger.log('Invalid JSON data format: ' + JSON.stringify(jsonData));
            throw new Error('Invalid JSON format. Expected object with "slides" array or an array of slides.');
        }

        // Flatten the payload so api.ts receives everything in one object
        const payload = {
            title: title,
            slides: slidesData,
            theme: theme,
            templateId: templateId || undefined,
            destinationId: destinationId, // Pass the pre-copied ID
            settings: undefined // Initialize
        };

        // Attach custom color settings if available in jsonData or Properties
        const userColors = getUserColors();
        // Priority: JSON settings > User Properties > Defaults
        const mergedColors = { ...userColors, ...(jsonData.settings && jsonData.settings.colors ? jsonData.settings.colors : {}) };

        if (Object.keys(mergedColors).length > 0) {
            payload.settings = { colors: mergedColors };
        }

        console.log('Checking for SlideGeneratorApi library...');
        console.log('Type of SlideGeneratorApi:', typeof SlideGeneratorApi);

        // Call the Library Function directly
        // Note: User must add the library with the identifier 'SlideGeneratorApi'
        if (typeof SlideGeneratorApi === 'undefined') {
            console.error('SlideGeneratorApi is undefined. Library configuration missing.');
            throw new Error('SlideGeneratorApi library not found. Please add the library with ID provided in instructions.');
        }

        console.log('Calling SlideGeneratorApi.generateSlides...');
        console.log('Payload being sent to library:', JSON.stringify(payload)); // DEBUG LOG
        const result = SlideGeneratorApi.generateSlides(payload);
        console.log('Library Result:', JSON.stringify(result));

        if (!result.success) {
            throw new Error(`Slide Generator Library Error: ${result.error}`);
        }

        // Expected result format: { success: true, url: "..." }
        if (!result.url) {
            throw new Error('Invalid response from Library: URL not found');
        }

        return {
            success: true,
            slideUrl: result.url
        };

    } catch (error) { // Type assertion
        Logger.log('Error in generateSlides: ' + error.toString());
        throw error;
    }
}


// ========================================
// User Settings (Colors)
// ========================================

/**
 * Save user defined colors to Properties
 */
function saveUserColors(colors) {
    const props = PropertiesService.getUserProperties();
    props.setProperties({
        'COLOR_PRIMARY': colors.primary,
        'COLOR_DEEP': colors.deep,
        'COLOR_TEXT': colors.text
    });
}

/**
 * Get user defined colors or defaults
 */
function getUserColors() {
    const props = PropertiesService.getUserProperties();
    return {
        primary: props.getProperty('COLOR_PRIMARY') || '#8FB130',
        deep: props.getProperty('COLOR_DEEP') || '#526717',
        text: props.getProperty('COLOR_TEXT') || '#333333'
    };
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
        return getMockData();
    }
    // =============================

    const projectId = getGcpProjectId();
    const location = getVertexAiLocation();
    const model = getVertexAiModel();

    Logger.log(`[callVertexAI] Config: ProjectId=${projectId}, Location=${location}, Model=${model}, UseMock=${useMock}`);

    if (useMock) {
        Logger.log('[callVertexAI] Mock mode enabled. Returning mock data.');
        try {
            const data = getMockData();
            Logger.log('[callVertexAI] Mock data retrieved successfully.');
            return data;
        } catch (e) {
            Logger.log('[callVertexAI] Error retrieving mock data: ' + e.toString());
            throw new Error('Failed to retrieve mock data: ' + e.toString());
        }
    }

    if (!projectId || projectId === 'YOUR_GCP_PROJECT_ID') {
        Logger.log('[callVertexAI] GCP Project ID missing.');
        throw new Error('GCP Project ID not configured.');
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
あなたはプロのプレゼンテーションデザイナーです。以下のドキュメント内容を、Google Slides用の構造化JSON形式に変換してください。

ターゲット層: ビジネスプロフェッショナル
トーン: プロフェッショナル、明確、簡潔

ドキュメント内容:
---
${documentContent}
---

コンテンツの長さに応じて、5〜12枚のスライド構成を作成してください。
出力は以下の厳格なスキーマに従った有効なJSONオブジェクトでなければなりません:

{
  "theme": "modern", // 推奨テーマ (例: modern, corporate, creative)
  "metada": {
    "estimatedDuration": "10 min",
    "audience": "General"
  },
  "slides": [
    {
      "type": "TITLE",
      "title": "プレゼンテーションのタイトル",
      "subtitle": "サブタイトルまたは発表者名",
      "notes": "タイトルスライドのスピーカーノート"
    },
    {
      "type": "AGENDA",
      "title": "アジェンダ",
      "items": ["トピック1", "トピック2", "トピック3"]
    },
    {
      "type": "CONTENT",
      "title": "スライドのタイトル",
      "body": {
        "text": "本文テキスト（ある場合）",
        "bullets": ["キーポイント1", "キーポイント2", "キーポイント3"]
      },
      "notes": "内容を説明するスピーカーノート"
    },
    {
      "type": "SECTION",
      "title": "セクションタイトル",
      "subtitle": "セクションの簡単な説明"
    },
    {
      "type": "CONCLUSION",
      "title": "まとめ / 重要なポイント",
      "points": ["ポイント1", "ポイント2"]
    }
  ]
}

ルール:
1. **要約**: テキストをそのままコピーしないでください。長い段落は簡潔な箇条書きに要約してください。
2. **構造**: 論理的な流れ (タイトル -> アジェンダ -> コンテンツ -> まとめ) にしてください。
3. **コンテンツ**: 各スライドの箇条書きは読みやすさのため最大3〜5項目にしてください。
4. **JSONのみ**: 純粋なJSONを出力してください。マークダウン形式 (\`\`\`json) や冒頭のテキストは含めないでください。
5. **言語**: 出力は必ず日本語で行ってください。
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
        gemUrl: props.getProperty('GEM_URL') || '',
        slideGeneratorApiUrl: props.getProperty('SLIDE_GENERATOR_API_URL') || '',
        useMockResponse: userProps.getProperty('USE_MOCK_RESPONSE') || 'false'
    };
}


