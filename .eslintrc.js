require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/default'],
  parserOptions: { tsconfigRootDir: __dirname },
  overrides: [
    {
      files: ['*.ts', '*.tsx'],
      parser: '@typescript-eslint/parser',
      'parserOptions': {
        'project': './tsconfig.json',
        'ecmaVersion': 2018,
        'sourceType': 'module'
      },
      rules: {
        'semi': "off",
        'dot-notation': 'off',
        '@typescript-eslint/no-explicit-any': 'off',
        'prefer-const': 'off',
        '@typescript-eslint/typedef': 'off',
        'eqeqeq': 'off',
        '@typescript-eslint/naming-convention': 'off',
        '@typescript-eslint/member-ordering': 'off',
        '@typescript-eslint/no-floating-promises': 'off',
        '@typescript-eslint/explicit-function-return-type': 'off',
        '@microsoft/spfx/no-async-await': 'off',
        'no-var': 'off',
        'no-return-assign': 'off',
        'no-eval': 'off',
        '@typescript-eslint/no-unused-vars': 'off',
        '@microsoft/spfx/pair-react-dom-render-unmount': 'off',
        '@typescript-eslint/explicit-member-accessibility': 'off',
        '@rushstack/security/no-unsafe-regexp': 'off',
        'require-atomic-updates': 'off',
        'no-empty-pattern': 'off',
        'no-empty': 'off',
        'guard-for-in': 'off',
        '@typescript-eslint/no-for-in-array': 'off',
        'no-cond-assign': 'off',
        'no-constant-condition': 'off',
        // '@typescript-eslint/no-use-before-define': 'off',
        'no-bitwise': 'off',
        '@typescript-eslint/no-empty-interface': [
          'error',
          {
            'allowSingleExtends': false
          }
        ],
        '@typescript-eslint/ban-types': [
          'error',
          {
            'types': {
              '{}': false
            }
          }
        ],
        'no-useless-escape': 'off',
        '@typescript-eslint/no-empty-interface': 'off',
        '@typescript-eslint/no-inferrable-types': 'off',
        '@typescript-eslint/no-non-null-assertion': 'off',
        'no-case-declarations': 'off',
        '@typescript-eslint/no-empty-function': 'off'

      }
    },
    {
      // For unit tests, we can be a little bit less strict.  The settings below revise the
      // defaults specified in the extended configurations, as well as above.
      files: [
        // Test files
        '*.test.ts',
        '*.test.tsx',
        '*.spec.ts',
        '*.spec.tsx',

        // Facebook convention
        '**/__mocks__/*.ts',
        '**/__mocks__/*.tsx',
        '**/__tests__/*.ts',
        '**/__tests__/*.tsx',

        // Microsoft convention
        '**/test/*.ts',
        '**/test/*.tsx'
      ],
      rules: {}
    }
  ]
};