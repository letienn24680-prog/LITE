import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  base: './',
  plugins: [react()],
  build: {
    target: 'esnext',
    chunkSizeWarningLimit: 1200,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (!id.includes('node_modules')) {
            return undefined;
          }
          if (id.includes('/node_modules/xlsx/')) {
            return 'vendor-xlsx';
          }
          if (id.includes('/node_modules/exceljs/')) {
            return 'vendor-exceljs';
          }
          if (id.includes('/node_modules/pdfjs-dist/')) {
            return 'vendor-pdfjs';
          }
          if (id.includes('/node_modules/pdf-lib/')) {
            return 'vendor-pdflib';
          }
          if (id.includes('/node_modules/jszip/')) {
            return 'vendor-jszip';
          }
          if (id.includes('/node_modules/pptxgenjs/')) {
            return 'vendor-pptxgen';
          }
          if (
            id.includes('/node_modules/docx-preview/') ||
            id.includes('/node_modules/mammoth/')
          ) {
            return 'vendor-docx';
          }
          if (id.includes('/node_modules/html2canvas/')) {
            return 'vendor-html2canvas';
          }
          if (id.includes('/node_modules/lucide-react/')) {
            return 'vendor-lucide';
          }
          if (id.includes('/node_modules/@supabase/')) {
            return 'vendor-supabase';
          }
          if (
            id.includes('/node_modules/react/') ||
            id.includes('/node_modules/react-dom/') ||
            id.includes('/node_modules/scheduler/')
          ) {
            return 'vendor-react';
          }
          return undefined;
        },
      },
    },
  },
  optimizeDeps: {
    esbuildOptions: {
      target: 'esnext',
    },
  },
});
