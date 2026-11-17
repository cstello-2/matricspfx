import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';
import { spMock } from './sp-mock/plugin';
import { aiMock } from './ai-mock/plugin';

const here = fileURLToPath(new URL('.', import.meta.url));
const repoRoot = fileURLToPath(new URL('..', import.meta.url));

export default defineConfig({
  root: here,
  plugins: [
    react(),
    spMock({
      fixturesDir: fileURLToPath(new URL('./fixtures', import.meta.url)),
      currentUser: { id: 1, name: 'Dev User', email: 'dev.user@contoso.dev' },
    }),
    aiMock('/mock-ai/v1'),
  ],
  resolve: {
    dedupe: ['react', 'react-dom'],
  },
  server: {
    port: 4400,
    strictPort: false,
    fs: { allow: [repoRoot] },
  },
  optimizeDeps: {
    include: [
      'react',
      'react-dom/client',
      'echarts/core',
      'echarts/charts',
      'echarts/components',
      'echarts/renderers',
      '@fluentui/react/lib/Styling',
      '@fluentui/react/lib/Theme',
      '@fluentui/react/lib/ThemeGenerator',
      '@fluentui/react/lib/Button',
      '@fluentui/react/lib/Stack',
      '@fluentui/react/lib/Text',
      '@fluentui/react/lib/Panel',
      '@fluentui/react/lib/Pivot',
      '@fluentui/react/lib/Modal',
      '@fluentui/react/lib/Dropdown',
      '@fluentui/react/lib/DatePicker',
      '@fluentui/react/lib/TextField',
      '@fluentui/react/lib/Toggle',
      '@fluentui/react/lib/SearchBox',
      '@fluentui/react/lib/Persona',
      '@fluentui/react/lib/Spinner',
      '@fluentui/react/lib/MessageBar',
      '@fluentui/react/lib/Link',
      '@fluentui/react/lib/ChoiceGroup',
      '@fluentui/react/lib/ProgressIndicator',
      '@fluentui/react/lib/Icon',
      '@fluentui/react/lib/Shimmer',
      '@fluentui/react/lib/Dialog',
      '@fluentui/react/lib/Label',
      '@fluentui/react/lib/Pickers',
      '@fluentui/font-icons-mdl2',
      '@fluentui/scheme-utilities',
      'jsonc-parser',
      '@codemirror/view',
      '@codemirror/state',
      '@codemirror/commands',
      '@codemirror/lang-json',
      '@codemirror/lint',
      '@codemirror/autocomplete',
      '@codemirror/language',
      '@lezer/highlight',
    ],
  },
});
