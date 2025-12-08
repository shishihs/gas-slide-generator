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
        .addItem('Create Template', 'createTemplateSlide')
        .addItem('Settings', 'showSettingsSidebar')
        .addToUi();
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
        // 1. Create a new presentation and get ID
        Logger.log('Creating new presentation...');
        let presentation = SlidesApp.create('Slide Generator Template');
        const presentationId = presentation.getId();
        const presentationUrl = presentation.getUrl();

        // Save and close to ensure creation is finalized
        presentation.saveAndClose();
        Logger.log('Presentation created and closed. ID: ' + presentationId);

        // 2. Re-open the presentation to modify it
        // This 'create -> close -> open' pattern helps avoid race conditions with new files
        Utilities.sleep(1000); // Wait a second for Drive propagation
        presentation = SlidesApp.openById(presentationId);
        Logger.log('Presentation re-opened.');

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

        // Show success message with instructions
        ui.alert(
            '✅ テンプレート作成完了',
            `テンプレートスライドを作成し、設定に保存しました。\n` +
            `作成予定: ${layoutConfigs.length} 枚\n` +
            `作成成功: ${createdCount} 枚\n\n` +
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
            `URL: ${presentationUrl}`,
            ui.ButtonSet.OK
        );

        Logger.log('Template created: ' + presentationId);

    } catch (error) {
        ui.alert('エラー', 'テンプレート作成に失敗しました: ' + error.toString(), ui.ButtonSet.OK);
        Logger.log('Error creating template: ' + error.toString());
    }
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
 * Generate slides from JSON data using the external Slide Generator API Library
 * Called from the sidebar
 * 
 * @param {Object} jsonData - The JSON data containing slide information
 * @returns {Object} - Result containing the new slide URL
 */
function generateSlidesFromJson(jsonData) {
    try {
        const templateId = getTemplateSlideId();

        // Check if jsonData is an array (from new Mock structure) or invalid
        let slidesData = [];
        let title = 'Generated Presentation';
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
            templateId: templateId || undefined
        };

        console.log('Checking for SlideGeneratorApi library...');
        console.log('Type of SlideGeneratorApi:', typeof SlideGeneratorApi);

        // Call the Library Function directly
        // Note: User must add the library with the identifier 'SlideGeneratorApi'
        if (typeof SlideGeneratorApi === 'undefined') {
            console.error('SlideGeneratorApi is undefined. Library configuration missing.');
            throw new Error('SlideGeneratorApi library not found. Please add the library with ID provided in instructions.');
        }

        console.log('Calling SlideGeneratorApi.generateSlides...');
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
        Logger.log('Using Mock Vertex AI Response - Comprehensive Test Data');
        // Comprehensive test data covering all slide types and chart types
        return [
            { "type": "title", "title": "包括的動作確認テスト用スライドデータ", "date": "2025.12.09", "notes": "これはタイトルスライドのスピーカーノートです。プレーンテキストのみ使用可能です。" },
            { "type": "agenda", "title": "本日のアジェンダ", "subhead": "全機能の動作確認を行います", "items": ["基本スライドタイプの確認", "比較系パターンのテスト", "プロセス・タイムライン系パターン", "データビジュアライゼーション", "グラフ描画機能の検証"], "notes": "本日は5つの章に分けて、すべてのスライドパターンとグラフ機能をテストします。" },
            { "type": "section", "title": "第1章：基本スライドタイプ", "sectionNo": 1, "notes": "まずは基本的なスライドタイプから確認していきましょう。" },
            { "type": "content", "title": "1カラムコンテンツスライド", "subhead": "基本的な箇条書きレイアウト", "points": ["第1ポイント：基本的なテキスト表示", "第2ポイント：**太字テキスト**のテスト", "第3ポイント：[[重要語]]の強調表示テスト", "第4ポイント：**太字**と[[強調語]]の組み合わせ"], "notes": "1カラムレイアウトの基本的な使い方を説明します。" },
            { "type": "content", "title": "2カラムコンテンツスライド", "subhead": "左右に分割されたレイアウト", "twoColumn": true, "columns": [["左カラム項目1", "左カラム項目2", "左カラム項目3"], ["右カラム項目1", "右カラム項目2", "右カラム項目3"]], "notes": "2カラムレイアウトでは、情報を左右に分けて表示できます。" },
            { "type": "cards", "title": "シンプルカード（文字列配列）", "subhead": "最もシンプルなカード表示", "columns": 3, "items": ["カードA", "カードB", "カードC", "カードD", "カードE", "カードF"], "notes": "6枚のシンプルカードを3列×2行で表示します。" },
            { "type": "headerCards", "title": "ヘッダー付きカード（3列）", "subhead": "色付きヘッダーのカード表示", "columns": 3, "items": [{ "title": "ヘッダー1", "desc": "カード1の説明文。ヘッダーは白文字で表示されます" }, { "title": "ヘッダー2", "desc": "カード2の説明文" }, { "title": "ヘッダー3", "desc": "カード3の説明文" }], "notes": "ヘッダーカードは色付きの背景に白文字で表示されます。" },
            { "type": "bulletCards", "title": "箇条書きカード（最大3枚）", "subhead": "詳細な説明を含むカード", "items": [{ "title": "ポイント1", "desc": "第1ポイントの詳細な説明文。[[重要な部分]]を強調できます" }, { "title": "ポイント2", "desc": "第2ポイントの詳細な説明文。**太字**も使用可能" }, { "title": "ポイント3", "desc": "第3ポイントの詳細な説明文" }], "notes": "箇条書きカードは最大3枚まで配置可能です。" },
            { "type": "table", "title": "テーブルスライド", "subhead": "表形式でのデータ表示", "headers": ["項目", "値", "備考"], "rows": [["項目A", "100", "正常"], ["項目B", "200", "注意"], ["項目C", "300", "警告"]], "notes": "テーブル形式でデータを整理して表示します。" },
            { "type": "section", "title": "第2章：比較系パターン", "sectionNo": 2 },
            { "type": "compare", "title": "左右比較スライド", "subhead": "2つの選択肢を比較", "leftTitle": "プランA", "rightTitle": "プランB", "leftItems": ["コスト効率が高い", "導入が容易", "サポート体制が充実"], "rightItems": ["機能が豊富", "拡張性が高い", "カスタマイズ可能"], "notes": "プランAとプランBを直接比較しています。" },
            { "type": "statsCompare", "title": "数値比較（テーブル形式）", "subhead": "Before/Afterの数値比較", "leftTitle": "導入前", "rightTitle": "導入後", "stats": [{ "label": "処理時間", "leftValue": "4時間", "rightValue": "30分", "trend": "down" }, { "label": "エラー率", "leftValue": "5.2%", "rightValue": "0.3%", "trend": "down" }, { "label": "満足度", "leftValue": "62%", "rightValue": "94%", "trend": "up" }], "notes": "導入前後の主要指標を比較しています。" },
            { "type": "barCompare", "title": "棒グラフ比較", "subhead": "視覚的な数値比較", "stats": [{ "label": "売上", "leftValue": "120", "rightValue": "150", "trend": "up" }, { "label": "利益", "leftValue": "30", "rightValue": "45", "trend": "up" }], "showTrends": true, "notes": "横棒グラフで視覚的に比較を表示します。" },
            { "type": "section", "title": "第3章：プロセス・タイムライン系", "sectionNo": 3 },
            { "type": "process", "title": "プロセススライド（視覚的形式）", "subhead": "最大4ステップの工程表示", "steps": ["要件定義", "設計", "実装", "テスト"], "notes": "ソフトウェア開発の基本的なプロセスを4ステップで示しています。" },
            { "type": "processList", "title": "プロセスリスト（リスト形式）", "subhead": "詳細な手順リスト", "steps": ["アカウントを作成する", "必要情報を入力する", "本人確認書類をアップロードする", "審査結果を待つ", "承認後、サービスを利用開始する"], "notes": "5つのステップでサービス利用開始までの流れを説明します。" },
            { "type": "timeline", "title": "タイムラインスライド", "subhead": "プロジェクトの進捗状況", "milestones": [{ "label": "企画立案完了", "date": "2024年1月", "state": "done" }, { "label": "設計フェーズ終了", "date": "2024年3月", "state": "done" }, { "label": "開発フェーズ", "date": "2024年6月", "state": "next" }, { "label": "リリース", "date": "2024年12月", "state": "todo" }], "notes": "現在、開発フェーズに入っています。" },
            { "type": "flowChart", "title": "フローチャート（2行）", "subhead": "複雑なフローを2行で表現", "flows": [{ "steps": ["企画", "設計", "開発", "テスト"] }, { "steps": ["デプロイ", "監視", "改善"] }] },
            { "type": "stepUp", "title": "ステップアップ（成長・進化）", "subhead": "段階的な成長を視覚化", "items": [{ "title": "初級", "desc": "基礎知識の習得" }, { "title": "中級", "desc": "応用力の向上" }, { "title": "上級", "desc": "専門性の確立" }, { "title": "マスター", "desc": "業界トップレベル" }], "notes": "4段階のスキルレベルを階段状に表示します。" },
            { "type": "section", "title": "第4章：図解系パターン", "sectionNo": 4 },
            { "type": "diagram", "title": "レーン図（スイムレーン）", "subhead": "役割別の作業分担を表示", "lanes": [{ "title": "営業部", "items": ["顧客訪問", "提案書作成", "契約締結"] }, { "title": "開発部", "items": ["要件定義", "システム開発", "テスト実施"] }], "notes": "2つの部署それぞれの担当業務を示しています。" },
            { "type": "cycle", "title": "サイクル図（4項目固定）", "subhead": "PDCAサイクルの例", "items": [{ "label": "Plan", "subLabel": "計画" }, { "label": "Do", "subLabel": "実行" }, { "label": "Check", "subLabel": "評価" }, { "label": "Act", "subLabel": "改善" }], "centerText": "PDCA", "notes": "PDCAサイクルを循環図で表現しています。" },
            { "type": "triangle", "title": "トライアングル図（3項目固定）", "subhead": "3つの要素の関係性を視覚化", "items": [{ "title": "品質", "desc": "高品質な成果物" }, { "title": "コスト", "desc": "予算内での実現" }, { "title": "納期", "desc": "期限内の完了" }], "notes": "プロジェクト管理の3つの制約を三角形で表現しています。" },
            { "type": "pyramid", "title": "ピラミッド図（階層構造）", "subhead": "マズローの欲求階層説", "levels": [{ "title": "自己実現", "description": "自己の可能性を最大限に発揮" }, { "title": "承認欲求", "description": "他者からの評価や尊敬" }, { "title": "社会的欲求", "description": "所属や愛情への欲求" }], "notes": "人間の欲求を3段階のピラミッドで表現しています。" },
            { "type": "section", "title": "第5章：データ・指標系パターン", "sectionNo": 5 },
            { "type": "kpi", "title": "KPIダッシュボード", "subhead": "主要業績評価指標", "columns": 4, "items": [{ "label": "売上高", "value": "12.5億円", "change": "+15%", "status": "good" }, { "label": "利益率", "value": "8.2%", "change": "-2%", "status": "bad" }, { "label": "顧客数", "value": "1,250", "change": "+120", "status": "good" }, { "label": "解約率", "value": "2.1%", "change": "±0%", "status": "neutral" }], "notes": "4つの主要KPIを表示しています。" },
            { "type": "progress", "title": "進捗状況", "subhead": "プロジェクト各フェーズの進捗", "items": [{ "label": "要件定義", "percent": 100 }, { "label": "設計", "percent": 100 }, { "label": "開発", "percent": 75 }, { "label": "テスト", "percent": 30 }], "notes": "開発フェーズは75%完了。" },
            { "type": "quote", "title": "引用スライド", "subhead": "著名人の言葉", "text": "Stay hungry, stay foolish.", "author": "Steve Jobs", "notes": "スティーブ・ジョブズの有名なスピーチからの引用です。" },
            { "type": "faq", "title": "よくある質問", "subhead": "お客様からのご質問", "items": [{ "q": "料金はいくらですか？", "a": "月額5,000円からご利用いただけます" }, { "q": "無料トライアルはありますか？", "a": "14日間の無料トライアルをご用意しています" }, { "q": "解約は簡単にできますか？", "a": "いつでもオンラインで即時解約可能です" }], "notes": "よくある3つの質問とその回答を掲載しています。" },
            { "type": "imageText", "title": "チャート表示テスト（折れ線）", "subhead": "グラフを画像として表示", "image": { "chartType": "line", "data": { "title": "月間売上推移", "subtitle": "（2024年度）", "source": "社内データ", "yAxisUnitLabel": "（万円）", "items": [{ "label": "4月", "value": 120 }, { "label": "5月", "value": 135 }, { "label": "6月", "value": 150 }, { "label": "7月", "value": 145 }, { "label": "8月", "value": 160 }], "color": { "start": "#e68a9c", "end": "#b469b8", "line": "#b469b8", "label": "#8c4fc8" }, "layout": { "width": 600, "height": 465, "marginTop": 100, "marginBottom": 85, "marginLeft": 75, "marginRight": 25 }, "yAxis": { "max": 200, "min": 0, "tickCount": 4 }, "lineOptions": { "markerRadius": 5, "dataLabelOffsetY": 15, "horizontalPadding": 30 } } }, "points": ["単一系列の折れ線グラフ", "グラデーション付きエリア表示", "データラベル表示対応"], "notes": "月間売上の推移を折れ線グラフで表示しています。" },
            { "type": "imageText", "title": "ドーナツグラフ", "image": { "chartType": "donut", "data": { "title": "市場シェア", "subtitle": "2024年度", "source": "業界レポート", "centerLabel": "シェア", "colors": [{ "id": "A", "start": "#e68a9c", "end": "#d96d8f" }, { "id": "B", "start": "#b469b8", "end": "#a656ad" }, { "id": "C", "start": "#9f63d0", "end": "#8c4fc8" }, { "id": "D", "start": "#7c6ce8", "end": "#6b5ce0" }], "items": [{ "label": "A社", "value": 35, "id": "A" }, { "label": "B社", "value": 25, "id": "B" }, { "label": "C社", "value": 20, "id": "C" }, { "label": "その他", "value": 20, "id": "D" }] } }, "points": ["円グラフの変形版", "中央に合計値を表示", "最大5項目を推奨"] },
            { "type": "closing", "notes": "ご清聴ありがとうございました。ご質問があればお気軽にどうぞ。" }
        ];
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
