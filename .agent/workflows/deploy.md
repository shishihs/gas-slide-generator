---
description: Build and Deploy the Slide Generator API and Add-on

---

## Slide Generator API

1. Build API
```bash
cd apps/slide-generator-api && npm install && npm run build
```

2. Push API to GAS
```bash
// turbo
cd apps/slide-generator-api && npx clasp push
```

## Doc to Slide Add-on

3. Build Add-on (generates HelpSidebar.html from docs)
```bash
cd apps/doc-to-slide-addon && npm run build
```

4. Push Add-on to GAS
```bash
// turbo
cd apps/doc-to-slide-addon && npx clasp push
```
