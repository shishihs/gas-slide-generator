import { defineConfig } from 'vite';
import path from 'path';

export default defineConfig({
    build: {
        lib: {
            entry: path.resolve(__dirname, 'src/api.ts'),
            name: 'SlideGeneratorApi',
            fileName: () => 'api.gs',
            formats: ['iife'],
        },
        outDir: 'dist',
        emptyOutDir: false, // appsscript.json をコピーするスクリプトが別途走るならfalseにするか、viteプラグインでコピーするか。ここでは手動コピーを想定してfalseにするか、あるいはvite-static-copyプラグインを使うのがスマート。
        // 今回は package.json の scripts で `mkdir -p dist && cp ...` しているので、emptyOutDir: true だと消される可能性がある。
        // しかしvite build が走ると自動で消しに来る。
        // script: "mkdir -p dist && cp src/appsscript.json dist/ && vite build" の順なら、vite build が dist を掃除してからビルドするので、appsscript.json が消える。
        // なので "vite build && cp src/appsscript.json dist/" の順にするのが正解。

        rollupOptions: {
            output: {
                extend: true,
                banner: 'var global = this;',
                footer: `
function generateSlides(data){ return global.generateSlides(data); }
function doPost(e){ return global.doPost(e); }
function doGet(e){ return global.doGet(e); }
`
            }
        },
        minify: false, // 読みやすさ重視
        target: 'es2019', // GAS compatible (Apps Script runtime is V8, ES2019 is safe)
    },
});
