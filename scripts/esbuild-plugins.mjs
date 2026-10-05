import { readFile } from "fs/promises";

/**
 * ExcelJS's browser bundle ships old setImmediate/microtask polyfills that, on
 * Internet Explorer, create empty <script> elements to schedule callbacks.
 * Obsidian never takes that path (MutationObserver/MessageChannel exist), but
 * script creation is not allowed in community plugins, so the conditions are
 * replaced with `false` and the minifier drops those branches.
 */
const SCRIPT_POLYFILLS = [
	/"document"in (\w+)&&"onreadystatechange"in \1\.document\.createElement\("script"\)/g,
	/(\w+)&&"onreadystatechange"in \1\.createElement\("script"\)/g,
	/"onreadystatechange"in (\w+)\("script"\)/g,
];

export const stripScriptPolyfills = {
	name: "strip-script-polyfills",
	setup(build) {
		build.onLoad({ filter: /exceljs[\\/]dist[\\/]exceljs(\.min)?\.js$/ }, async (args) => {
			let code = await readFile(args.path, "utf8");
			for (const re of SCRIPT_POLYFILLS) {
				const before = code;
				code = code.replace(re, "!1");
				if (code === before) {
					throw new Error(`strip-script-polyfills: pattern ${re} not found in ${args.path} (did ExcelJS change?)`);
				}
			}
			return { contents: code, loader: "js" };
		});
	},
};

/** Fails the production build if the bundle can still create <script> elements. */
export async function assertNoScriptCreation(file) {
	const code = await readFile(file, "utf8");
	const hits = code.match(/createElement\(\s*["'`]script["'`]\s*\)|\(\s*["'`]script["'`]\s*\)\.onreadystatechange/g);
	if (hits) {
		throw new Error(`${file} creates <script> elements (${hits.length}x): not allowed in Obsidian community plugins.`);
	}
}
