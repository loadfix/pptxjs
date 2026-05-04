import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';

const output = {
	banner: `/*
 * @license
 * pptx-preview
 * Released under Apache License 2.0
 */`,
	sourcemap: true,
}

const umdOutput = {
	...output,
	name: "pptx",
	file: 'dist/pptx-preview.js',
	format: 'umd',
	globals: {
		jszip: 'JSZip'
	},
};

export default args => {
	const config = {
		input: 'src/pptx-preview.ts',
		output: [umdOutput],
		external: ['jszip'],
		plugins: [typescript()]
	}

	if (args.environment == 'BUILD:production')
		config.output = [umdOutput,
			{
				...umdOutput,
				file: 'dist/pptx-preview.min.js',
				plugins: [terser()]
			},
			{
				...output,
				file: 'dist/pptx-preview.mjs',
				format: 'es',
			},
			{
				...output,
				file: 'dist/pptx-preview.min.mjs',
				format: 'es',
				plugins: [terser()]
			}];

	return config
};
