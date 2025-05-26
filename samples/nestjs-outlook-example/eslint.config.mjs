import tseslint from 'typescript-eslint';
import eslintPluginComments from '@eslint-community/eslint-plugin-eslint-comments';

export default tseslint.config(
  tseslint.configs.strictTypeChecked,
  {
    ignores: ['*.mjs', '*.cjs', '*.js', 'dist/**'],
    plugins: {
      'eslint-comments': eslintPluginComments,
    },
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
      
      // Require descriptions when using ESLint disable comments
      'eslint-comments/require-description': ['error', { ignore: [] }],

      // Disallow duplicate disable directives
      'eslint-comments/no-duplicate-disable': 'error',

      // Disallow unused disable directives
      'eslint-comments/no-unused-disable': 'error',

      // Prevent disabling all rules without specifics
      'eslint-comments/no-unlimited-disable': 'error',

      // Allow empty classes with decorators (for NestJS modules)
      '@typescript-eslint/no-extraneous-class': ['error', {
        allowWithDecorator: true
      }],

      // Relax this rule to allow number, boolean, and nullable
      '@typescript-eslint/restrict-template-expressions': ['error', {
        allow: [{ name: ['Error', 'URL', 'URLSearchParams'], from: 'lib' }],
        allowAny: true,
        allowBoolean: true,
        allowNullish: true,
        allowNumber: true,
        allowRegExp: true,
      }],
    },
    linterOptions: {
      reportUnusedDisableDirectives: true,
    },
  }
);