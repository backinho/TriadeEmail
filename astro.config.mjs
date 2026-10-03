import { defineConfig } from 'astro/config';

export default defineConfig({
  outDir: './web-dist',
  server: {
    host: true,
    port: 4321,
  },
  vite: {
    server: {
      watch: {
        ignored: ['**/release/**'],
      },
    },
  },
  devToolbar: {
    enabled: false,
  },
});
