const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  { ignores: ['tests/**', 'workbench/**'] },
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    }
  },
  {
    files: ['src/core/**', 'src/components/**'],
    rules: {
      '@rushstack/no-new-null': 'off'
    }
  }
];
