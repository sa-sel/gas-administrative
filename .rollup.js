import { babel } from '@rollup/plugin-babel';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import { resolve } from 'path';
import html from 'rollup-plugin-html';
import typescript from 'rollup-plugin-typescript2';

const extensions = ['.ts', '.js'];

const preventTreeShakingPlugin = {
  name: 'no-threeshaking',
  resolveId: (id, importer) => (importer ? null : { id, moduleSideEffects: 'no-treeshake' }),
};

const isGithub = process.env.ENV === 'GITHUB';
const minifiedJs = {
  comments: false,
  compact: true,
  minified: true,
};
const minifiedHtml = {
  collapseWhitespace: true,
  collapseBooleanAttributes: true,
  conservativeCollapse: true,
  minifyJS: true,
};

/** @type {import('rollup').RollupOptions} */
const config = {
  input: './.build/index.ts',
  output: { dir: '.build', format: 'esm' },
  plugins: [
    preventTreeShakingPlugin,
    nodeResolve({ extensions }),
    html({
      include: ['./src/views/*.html', './.build/views/*.html'],
      ...(isGithub ? minifiedHtml : {}),
    }),
    typescript(),
    babel({
      extensions,
      babelHelpers: 'runtime',
      include: resolve('.build', 'src', '**', '*.ts'),
      exclude: ['node_modules/**'],
      configFile: './.babelrc.js',
      ...(isGithub ? minifiedJs : {}),
    }),
  ],
};

export default config;
