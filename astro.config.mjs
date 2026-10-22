import { defineConfig } from 'astro/config';

export default defineConfig({
  site: 'https://delphine-code.github.io',
  base: '/jiashi-xinmei-site',
  trailingSlash: 'always',
  build: {
    inlineStylesheets: 'auto'
  },
  vite: {
    build: { chunkSizeWarningLimit: 1500 },
    optimizeDeps: { disabled: true }
  }
});
