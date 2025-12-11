
/**
 * Test Mock Data for Slide Generator
 * 
 * This file contains mock data used for testing the slide generation capabilities.
 * It provides examples for all supported slide types and patterns.
 */
function getMockData() {
    return [
        // === STRUCTURE SLIDES ===
        { "type": "title", "title": "[TEST] title タイプ", "subtitle": "全パターン検証用テストデータ", "date": "2025.12.11", "notes": "タイトルスライドのテスト" },

        { "type": "agenda", "title": "[TEST] agenda タイプ", "subhead": "目次・アジェンダ形式", "items": ["項目1: 基本機能", "項目2: 拡張機能", "項目3: 検証項目"], "notes": "アジェンダスライドのテスト" },

        { "type": "section", "title": "[TEST] section タイプ", "sectionNo": 1, "notes": "セクションスライドのテスト" },

        // === CONTENT SLIDES ===
        { "type": "content", "title": "[TEST] content タイプ (通常)", "subhead": "箇条書きコンテンツ", "points": ["ポイント1: 通常のテキスト", "ポイント2: 複数行のテキスト\\n改行を含む内容", "ポイント3: 最後の項目"], "notes": "通常コンテンツスライドのテスト" },

        // === CARD TYPES ===
        {
            "type": "cards", "title": "[TEST] cards タイプ (文字列配列)", "subhead": "items が文字列の場合", "columns": 3, "items": [
                "カード1タイトル\\n説明文がここに入ります。複数行対応。",
                "カード2タイトル\\n別の説明文です。",
                "カード3タイトル\\n3つ目の説明。"
            ], "notes": "cards (文字列配列) のテスト"
        },

        {
            "type": "cards", "title": "[TEST] cards タイプ (オブジェクト配列)", "subhead": "items がオブジェクトの場合", "columns": 2, "items": [
                { "title": "機能A", "desc": "機能Aの詳細説明。オブジェクト形式での記述。" },
                { "title": "機能B", "desc": "機能Bの詳細説明。" }
            ], "notes": "cards (オブジェクト配列) のテスト"
        },

        {
            "type": "headerCards", "title": "[TEST] headerCards タイプ", "subhead": "ヘッダー付きカード", "columns": 3, "items": [
                { "title": "ヘッダー1", "desc": "ヘッダー付きカードの本文。色付きヘッダーが表示されるべき。" },
                { "title": "ヘッダー2", "desc": "2枚目のカード本文。" },
                { "title": "ヘッダー3", "desc": "3枚目のカード本文。" }
            ], "notes": "headerCards のテスト"
        },

        // === KPI ===
        {
            "type": "kpi", "title": "[TEST] kpi タイプ", "subhead": "KPIダッシュボード", "items": [
                { "label": "売上", "value": "¥1.2億", "change": "+15%", "status": "good" },
                { "label": "コスト", "value": "¥3,500万", "change": "-5%", "status": "good" },
                { "label": "顧客数", "value": "1,234", "change": "-2%", "status": "bad" },
                { "label": "NPS", "value": "72", "change": "+8pt", "status": "good" }
            ], "notes": "KPIカードのテスト。good/bad ステータスで色が変わるべき。"
        },

        // === PROGRESS ===
        {
            "type": "progress", "title": "[TEST] progress タイプ", "subhead": "進捗バー表示", "items": [
                { "label": "設計フェーズ", "percent": 100 },
                { "label": "開発フェーズ", "percent": 75 },
                { "label": "テストフェーズ", "percent": 30 },
                { "label": "リリース", "percent": 0 }
            ], "notes": "プログレスバーのテスト。%表示が正しいか確認。"
        },

        // === TABLE ===
        {
            "type": "table", "title": "[TEST] table タイプ", "subhead": "表形式データ", "headers": ["項目", "値", "備考"], "rows": [
                ["データ1", "100", "正常"],
                ["データ2", "200", "注意"],
                ["データ3", "300", "完了"]
            ], "notes": "テーブルのテスト。ヘッダー行と交互の背景色を確認。"
        },

        // === FAQ ===
        {
            "type": "faq", "title": "[TEST] faq タイプ", "subhead": "Q&A形式", "items": [
                { "q": "質問1: FAQスライドはどう表示されますか？", "a": "回答1: Qマークと質問・回答がセットで表示されます。" },
                { "q": "質問2: 複数のQ&Aは対応していますか？", "a": "回答2: はい、複数表示されます。" }
            ], "notes": "FAQスライドのテスト。Q/A の視覚的区別を確認。"
        },

        // === QUOTE ===
        { "type": "quote", "title": "[TEST] quote タイプ", "subhead": "引用・メッセージ", "text": "これはテスト用の引用文です。大きく中央揃えで表示されるべきです。", "author": "テスト担当者", "notes": "引用スライドのテスト。引用符と著者名を確認。" },

        // === TIMELINE ===
        {
            "type": "timeline", "title": "[TEST] timeline タイプ", "subhead": "時系列表示", "milestones": [
                { "label": "過去イベント", "date": "2024年", "state": "past" },
                { "label": "現在進行中", "date": "2025年", "state": "current" },
                { "label": "将来の計画", "date": "2026年", "state": "future" }
            ], "notes": "タイムラインのテスト。past/current/future の状態表示を確認。"
        },

        // === PROCESS ===
        {
            "type": "process", "title": "[TEST] process タイプ", "subhead": "プロセスフロー", "steps": [
                "Step 1\\n最初のステップ",
                "Step 2\\n2番目のステップ",
                "Step 3\\n3番目のステップ",
                "Step 4\\n最後のステップ"
            ], "notes": "プロセスフローのテスト。矢印の接続を確認。"
        },

        // === FLOWCHART ===
        {
            "type": "flowChart", "title": "[TEST] flowChart タイプ", "subhead": "フローチャート形式", "steps": [
                "開始",
                "処理A",
                "処理B",
                "終了"
            ], "notes": "フローチャートのテスト。ボックスと矢印の表示を確認。"
        },

        // === STEPUP ===
        {
            "type": "stepUp", "title": "[TEST] stepUp タイプ", "subhead": "ステップアップ (階段表示)", "items": [
                { "title": "Level 1", "desc": "基礎レベル" },
                { "title": "Level 2", "desc": "中級レベル" },
                { "title": "Level 3", "desc": "上級レベル" }
            ], "notes": "ステップアップのテスト。階段状の表示を確認。"
        },

        // === CYCLE ===
        {
            "type": "cycle", "title": "[TEST] cycle タイプ", "subhead": "循環フロー", "items": [
                { "label": "計画", "subLabel": "Plan" },
                { "label": "実行", "subLabel": "Do" },
                { "label": "評価", "subLabel": "Check" },
                { "label": "改善", "subLabel": "Act" }
            ], "centerText": "PDCA", "notes": "サイクル図のテスト。円形配置と中心テキストを確認。"
        },

        // === TRIANGLE ===
        {
            "type": "triangle", "title": "[TEST] triangle タイプ", "subhead": "三角形構造", "items": [
                { "title": "頂点A", "desc": "説明A" },
                { "title": "頂点B", "desc": "説明B" },
                { "title": "頂点C", "desc": "説明C" }
            ], "notes": "三角形図のテスト。3つの頂点が正しく配置されるか確認。"
        },

        // === PYRAMID ===
        {
            "type": "pyramid", "title": "[TEST] pyramid タイプ", "subhead": "ピラミッド構造", "levels": [
                { "title": "トップ層", "description": "最上位の説明" },
                { "title": "中間層", "description": "中間層の説明" },
                { "title": "ベース層", "description": "基盤となる説明" }
            ], "notes": "ピラミッド図のテスト。上から下への階層表示を確認。"
        },

        // === DIAGRAM (LANES) ===
        {
            "type": "diagram", "title": "[TEST] diagram タイプ (レーン)", "subhead": "レーン形式の図", "lanes": [
                { "title": "レーン1", "items": ["項目A", "項目B", "項目C"] },
                { "title": "レーン2", "items": ["項目X", "項目Y"] },
                { "title": "レーン3", "items": ["単一項目"] }
            ], "notes": "レーン図のテスト。複数レーンと項目の表示を確認。"
        },

        // === COMPARE ===
        { "type": "compare", "title": "[TEST] compare タイプ (基本比較)", "subhead": "左右比較レイアウト", "leftTitle": "オプションA", "rightTitle": "オプションB", "leftItems": ["A の特徴1", "A の特徴2", "A の利点"], "rightItems": ["B の特徴1", "B の特徴2", "B の利点"], "notes": "基本比較スライドのテスト。左右のタイトルと項目を確認。" },

        // === STATSCOMPARE ===
        {
            "type": "statsCompare", "title": "[TEST] statsCompare タイプ", "subhead": "数値比較 (統計)", "leftTitle": "Before", "rightTitle": "After", "stats": [
                { "label": "指標1", "leftValue": "100", "rightValue": "150", "trend": "up" },
                { "label": "指標2", "leftValue": "50", "rightValue": "30", "trend": "down" }
            ], "notes": "統計比較のテスト。up/down の矢印表示を確認。"
        },

        // === BARCOMPARE ===
        {
            "type": "barCompare", "title": "[TEST] barCompare タイプ", "subhead": "バーチャート比較", "leftTitle": "カテゴリA", "rightTitle": "カテゴリB", "stats": [
                { "label": "項目1", "leftValue": "40", "rightValue": "80" },
                { "label": "項目2", "leftValue": "70", "rightValue": "50" }
            ], "notes": "バー比較のテスト。バーの長さが値に比例しているか確認。"
        },

        // === IMAGETEXT ===
        { "type": "imageText", "title": "[TEST] imageText タイプ", "subhead": "画像とテキスト", "image": "https://via.placeholder.com/400x300.png?text=Test+Image", "imagePosition": "left", "points": ["画像の右側にテキストが表示されるべき", "複数のポイントを記載可能"], "notes": "imageText のテスト。画像プレースホルダーとテキストの配置を確認。" },

        // === CLOSING ===
        { "type": "closing", "title": "[TEST] closing タイプ", "notes": "クロージングスライドのテスト。" }
    ];
}
