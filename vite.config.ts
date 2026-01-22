
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  define: {
    // لضمان وصول مفتاح API الخاص بـ Gemini إذا تم استخدامه
    'process.env': process.env
  },
  server: {
    port: 3000,
  },
  build: {
    outDir: 'dist',
    target: 'esnext'
  }
});
