module.exports = {
  ...require('gts/.prettierrc.json'),
  overrides: [
    {
      files: '*.html',
      options: {
        parser: 'angular',
        printWidth: 100,
      },
    },
  ],
};
