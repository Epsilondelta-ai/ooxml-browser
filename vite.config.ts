import { fileURLToPath, URL } from 'node:url';

import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    sourcemap: true
  },
  resolve: {
    alias: {
      '@ooxml/core': fileURLToPath(new URL('./packages/core/src/index.ts', import.meta.url)),
      '@ooxml/docx': fileURLToPath(new URL('./packages/docx/src/index.ts', import.meta.url)),
      '@ooxml/xlsx': fileURLToPath(new URL('./packages/xlsx/src/index.ts', import.meta.url)),
      '@ooxml/pptx': fileURLToPath(new URL('./packages/pptx/src/index.ts', import.meta.url)),
      '@ooxml/render': fileURLToPath(new URL('./packages/render/src/index.ts', import.meta.url)),
      '@ooxml/editor': fileURLToPath(new URL('./packages/editor/src/index.ts', import.meta.url)),
      '@ooxml/browser': fileURLToPath(new URL('./packages/browser/src/index.ts', import.meta.url)),
      '@ooxml/devtools': fileURLToPath(new URL('./packages/devtools/src/index.ts', import.meta.url))
    }
  },
  test: {
    environment: 'node'
  }
});
