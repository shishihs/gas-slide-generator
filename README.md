# Google Document to Slide Generator

📄 → 🤖 → 🎨

Google ドキュメントの内容を Vertex AI で解析し、Google Slides を自動生成するアドオンです。

## 機能

| メニュー項目 | 説明 |
|-------------|------|
| **Convert to JSON** | ドキュメントを Vertex AI で解析し、スライド用 JSON に変換 |
| **Generate Slide** | JSON を基にテンプレートから Google Slides を自動生成 |

## アーキテクチャ

```
GAS アドオン  ──────→  Vertex AI (Gemini)
              直接呼び出し
              (OAuth認証)
```

**Cloud Functions 不要！** GAS から Vertex AI を直接呼び出します。

## クイックスタート

詳細は [docs/README.md](./docs/README.md) を参照してください。

## ファイル構成

```
.
├── apps/
│   └── gas-addon/        # GAS アドオンコード
│       ├── Code.js
│       ├── GenerateSlideSidebar.html
│       ├── SettingsSidebar.html
│       └── appsscript.json
└── docs/
    └── README.md         # 詳細ドキュメント
```
