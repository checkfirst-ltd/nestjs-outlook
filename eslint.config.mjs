import tseslint from 'typescript-eslint';
import eslintPluginComments from '@eslint-community/eslint-plugin-eslint-comments';
import eslintPluginImport from 'eslint-plugin-import';

export default tseslint.config(
  tseslint.configs.strictTypeChecked,
  {
    ignores: ['*.mjs', '*.cjs', '*.js', 'dist/**'],
    plugins: {
      'eslint-comments': eslintPluginComments,
      'import': eslintPluginImport,
    },
    languageOptions: {
      parserOptions: {
        project: './tsconfig.json',
        tsconfigRootDir: import.meta.dirname,
      },
    },
    settings: {
      'import/resolver': {
        typescript: {
          project: './tsconfig.json',
        },
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

      // Enforce relative imports for local modules
      'import/no-absolute-path': 'error',

      // Prevent imports from 'src/' at the beginning (force relative imports)
      'import/no-useless-path-segments': ['error', {
        noUselessIndex: true,
      }],
    },
    linterOptions: {
      reportUnusedDisableDirectives: true,
    },
  }
);