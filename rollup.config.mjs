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
		jszip: 'JSZip',
		// EMF / WMF conversion is optional — UMD consumers can load
		// `emf-converter` via a second <script> tag to expose window.EMFConverter
		// (or the adapter will silently skip metafile images). The base bundle
		// intentionally does not pull this in because most decks don't carry
		// any EMF/WMF and we don't want to inflate the baseline for them.
		'emf-converter': 'EMFConverter'
	},
};

export default args => {
	const config = {
		input: 'src/pptx-preview.ts',
		output: [umdOutput],
		external: ['jszip', 'emf-converter'],
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
