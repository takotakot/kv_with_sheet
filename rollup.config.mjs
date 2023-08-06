import typescript from 'rollup-plugin-typescript2';
import cleanup from 'rollup-plugin-cleanup';
import license from 'rollup-plugin-license';
import { fileURLToPath } from 'url';
import multi from '@rollup/plugin-multi-entry';

export default {
  input: ['src/doPost.ts', 'src/testKv.ts', 'src/testKvConfig.ts'],
  output: {
    dir: 'dist',
    format: 'esm',
  },
  plugins: [
    cleanup({ comments: 'none', extensions: ['.ts'] }),
    license({
      banner: {
        content: {
          file: fileURLToPath(new URL('license-header.txt', import.meta.url)),
        },
      },
    }),
    typescript(),
    multi(),
  ],
  context: 'this',
};
