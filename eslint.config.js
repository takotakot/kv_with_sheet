let customConfig = [];
let hasIgnoresFile = false;
try {
  require.resolve('./eslint.ignores.cjs');
  hasIgnoresFile = true;
} catch {
  // eslint.ignores.cjs doesn't exist
}

if (hasIgnoresFile) {
  const ignores = require('./eslint.ignores.cjs');
  customConfig = [{ignores}];
} else {
  try {
    require.resolve('./eslint.ignores.js');
    const ignores = require('./eslint.ignores.js');
    customConfig = [{ignores}];
  } catch {
    // eslint.ignores.js doesn't exist
  }
}

module.exports = [
  ...customConfig,
  ...require('gts'),
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.json'],
      },
    },
  },
];
