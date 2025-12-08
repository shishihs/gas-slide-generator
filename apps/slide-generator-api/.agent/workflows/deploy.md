---
description: Build and Deploy the Slide Generator API
---

# Build and Deploy Workflow

This workflow describes the process for building and deploying the `slide-generator-api`.

**CRITICAL RULE: Always ensure the build and deploy steps are successfully completed BEFORE reporting status to the user.**

## Steps

1.  **Build**
    Run the build script to bundle TypeScript files and prepare the `dist` directory.
    ```bash
    npm run build
    ```
    *Ensure this command exits with code 0.*

2.  **Deploy**
    Push the built files to Google Apps Script project using clasp.
    ```bash
    clasp push
    ```
    *Ensure the output confirms files were pushed (e.g., "Pushed X files."). If login is required, notify the user immediately.*

3.  **Report**
    Only after step 1 and 2 are successful, confirm to the user that the changes are live.
