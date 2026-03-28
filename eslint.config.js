import js from '@eslint/js';
import tseslint from '@typescript-eslint/eslint-plugin';
import tsparser from '@typescript-eslint/parser';

const sharedGlobals = {
  window: 'readonly',
  document: 'readonly',
  Blob: 'readonly',
  File: 'readonly',
  HTMLElement: 'readonly',
  console: 'readonly',
  performance: 'readonly'
};

export default [
  {
    ignores: ['**/dist/**', 'node_modules/**', '.omx/**']
  },
  js.configs.recommended,
  {
    files: ['**/*.ts'],
    languageOptions: {
      parser: tsparser,
      parserOptions: {
        project: ['./tsconfig.eslint.json'],
        sourceType: 'module'
      },
      globals: sharedGlobals
    },
    plugins: {
      '@typescript-eslint': tseslint
    },
    rules: {
      '@typescript-eslint/consistent-type-imports': 'error',
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      'no-undef': 'off'
    }
  },
  {
    files: ['**/*.mjs'],
    languageOptions: {
      globals: sharedGlobals
    }
  }
];
