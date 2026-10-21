import { defineConfig } from 'astro/config';

export default defineConfig({
  site: 'https://jiashi-xinmei.vercel.app',
  build: {
    inlineStylesheets: 'auto'
  },
  vite: {
    build: { chunkSizeWarningLimit: 1500 },
    optimizeDeps: { disabled: true }
  }
});
