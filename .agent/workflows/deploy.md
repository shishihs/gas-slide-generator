---
description: Build and Deploy the Slide Generator API and Add-on
---

## デプロイ手順

### 1. ライブラリ (slide-generator-api) をビルド・デプロイ

```bash
// turbo
cd apps/slide-generator-api && npm run build && cp dist/api.gs ../doc-to-slide-addon/src/api.gs
```

### 2. ライブラリをGASにpush & バージョン作成

```bash
// turbo
cd apps/slide-generator-api && npx clasp push && npx clasp version "デプロイ"
```

### 3. アドオンのappsscript.jsonのライブラリバージョンを更新

**重要:** Step 2で作成されたバージョン番号を確認し、`apps/doc-to-slide-addon/src/appsscript.json` の `dependencies.libraries[0].version` を更新すること。

### 4. アドオン (doc-to-slide-addon) をGASにpush

```bash
// turbo
cd apps/doc-to-slide-addon && npx clasp push --force
```

### 5. Git コミット・プッシュ (ルートディレクトリから実行)

```bash
cd /Users/shishi/Workspace/gas-slide-generator && git add . && git commit -m "deploy: Update library and addon" && git push
```

## チェックリスト

- [ ] ライブラリをビルド
- [ ] ライブラリをclasp push
- [ ] ライブラリのバージョンを作成 (clasp version)
- [ ] **appsscript.jsonのライブラリバージョンを更新**
- [ ] アドオンをclasp push
- [ ] **ルートディレクトリからgit add/commit/push**
