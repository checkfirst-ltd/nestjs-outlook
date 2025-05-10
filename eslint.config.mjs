import tseslint from 'typescript-eslint';

export default tseslint.config(
  tseslint.configs.strictTypeChecked,
  {
    ignores: ['*.mjs', '*.cjs', '*.js', 'dist/**'],
    languageOptions: {
      parserOptions: {
        project: './tsconfig.json',
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // Customize ignore patterns for unused variables (not in base config)
      '@typescript-eslint/no-unused-vars': ['error', {
        argsIgnorePattern: '^_',
        varsIgnorePattern: '^_',
        caughtErrorsIgnorePattern: '^_',
      }],
      // Ensure ts-comments are not misused
      '@typescript-eslint/ban-ts-comment': ['error', {
        'ts-ignore': true,
        'ts-expect-error': false, // allow this instead
      }],
      // Not included in strict config — prefer @ts-expect-error over @ts-ignore
      '@typescript-eslint/prefer-ts-expect-error': 'warn',
    },
    linterOptions: {
      reportUnusedDisableDirectives: true,
    },
  }
);