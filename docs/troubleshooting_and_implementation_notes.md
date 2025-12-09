# Slide Generator Debugging & Implementation Notes

## 1. Template Access & Auto-Healing Mechanism
One of the primary issues encountered was the **"Failed to access or copy template slide"** error.

*   **Problem**: The script often failed to access the user-provided Template ID due to permissions or the file not strictly belonging to the script's execution context.
*   **Solution**: Implemented an "Auto-Healing" strategy.
    *   If the provided Template ID fails (or is invalid), the system automatically creates a *new* presentation from scratch using `SlidesApp.create()`.
    *   **Crucial Fix**: Added a `Utilities.sleep(5000)` delay after creating/copying files. Google Drive API has a propagation delay, and accessing a file immediately after creation often results in a 404/500 server error.

## 2. Library Build & Deployment Process
A persistent issue was that code fixes were not reflecting in the actual execution, leading to recurring errors like `p.getPlaceholderType is not a function` even after "fixing" the code.

*   **Problem**: The Google Apps Script library logic is written in TypeScript (`src/api.ts`, etc.) but must be bundled into a single `.gs` file (`dist/api.gs`) for execution on GAS.
*   **Discovery**: Modifying TS files alone is NOT enough. The build command `npm run build` must be executed to generate the valid `dist/api.gs`.
*   **Workflow**: 
    1.  `npm run build` (Compiles TS -> JS/GS)
    2.  `clasp push` (Uploads `dist` contents)
    3.  `clasp version "..."` (Creates an immutable library version)
    4.  Update `appsscript.json` in the consumer Add-on to use the new library version.
    5.  `clasp push` & `clasp deploy` the Add-on.

## 3. `getPlaceholderType` Error
*   **Error**: `TypeError: p.getPlaceholderType is not a function`.
*   **Cause**: `SlidesApp.getPlaceholders()` returns `PageElement` objects which act like Placeholders but sometimes don't have the `getPlaceholderType()` method directly on the wrapper in certain contexts, or when the element is actually a generic Shape acting as a placeholder.
*   **Fix**: Implemented a safe check:
    ```typescript
    if (typeof p.getPlaceholderType === 'function') { ... } 
    try { const s = p.asShape(); if (typeof s.getPlaceholderType === 'function') ... }
    ```

## 4. Data Loss in API Layer
*   **Problem**: Rich data (graphs, tables, diagram steps) sent from the Add-on (UI) was missing when it reached the slide generators.
*   **Cause**: The API entry point (`api.ts`) explicitly mapped specific fields (`title`, `subtitle`, `content`) and ignored the rest.
*   **Fix**: Updated the API to spread **all** input properties (`...s`) into the internal data model. Also updated the `Slide` domain model's constructor to accept a `rawData: any` argument to carry this payload.

## 5. Diagram & Chart Implementation
*   **Discovery**: Although the specification defined slide types like `timeline`, `process`, `cycle`, etc., there was **no implementation code** to actually draw them. They were falling back to default text slides.
*   **Solution**: Created `GasDiagramSlideGenerator.ts`.
    *   Since GAS cannot easily render SVG strings directly, we implemented "Programmatic Drawing" using `SlidesApp` Shape APIs.
    *   **Timeline**: Drawn using `ShapeType.RECTANGLE` (line) and `ELLIPSE` (milestones).
    *   **Process**: Drawn using `ShapeType.CHEVRON`.
    *   **Chart/Graph**: (To be refined) Currently handled via `imageText` or requires a dedicated charting service (Charts API).

## 6. Layout Selection Logic
*   **Problem**: Slides often defaulted to "Title and Body" because layout lookup was case-sensitive and rigid.
*   **Fix**:
    *   Normalized all layout lookups to uppercase (`.toUpperCase()`).
    *   Added robust fallback searching (e.g., if "SECTION_HEADER" is missing, try "SECTION ONLY").
    *   Added logging to `GasSlideRepository` to print all available layout names in the template for debugging.

## Future Recommendations
*   **Charts**: Implementing native Google Sheets charts embedding or more complex shape-based charts for data visualization.
*   **Robustness**: Continue using safely typed fallbacks (like `primary_color` if config is missing) to prevent runtime crashes.

## 7. Invalid Hex Color String Error
*   **Error**: `Exception: 無効な 16 進数の文字列: 「」は有効な 16 進数のカラーコードではありません。` (Invalid hex string: "" is not a valid hex color code).
*   **Cause**: The `SlideConfig.ts` file had empty strings `''` for several color definitions (e.g., `background_gray`, `process_arrow`). When code attempted to use these values (e.g., `setSolidFill(CONFIG.COLORS.process_arrow)`), the empty string caused a GAS runtime error.
*   **Fix**: Populated all empty color strings in `SlideConfig.ts` with valid hex codes or implemented explicit fallbacks in the generator code (e.g., `|| '#CCCCCC'`).

## 8. Missing Logic for Diagram Slides
*   **Problem**: Users requested slide types like `timeline`, `process`, `cycle` specified in the JSON, but the system rendered them as basic text slides.
*   **Investigation**: It was discovered that while the JSON parser accepted these types, there was **no corresponding Generator class** implemented to handle them. The system defaulted to `GasContentSlideGenerator` properly but silently.
*   **Solution**:
    *   Created `GasDiagramSlideGenerator.ts` to implement specific drawing logic for each diagram type using `SlidesApp` Shapes (Circles, Chevrons, Connectors).
    *   Updated `GasSlideRepository.ts` to instantiate this generator and dispatch relevant slide types (e.g., `timeline`, `process`, `stepup`) to it.

