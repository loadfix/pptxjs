# pptxjs — project notes for Claude

Browser-side PPTX → HTML renderer. TypeScript, built with rollup, tested with Playwright. Follows the same structure as the sibling [docxjs](https://github.com/loadfix/docxjs) project.

## OOXML feature workflow (required before adding rendering for any new feature)

Every OOXML feature is defined by a manifest in the shared corpus
repository `loadfix/ooxml-reference-corpus` (sibling checkout at
`../ooxml-reference-corpus/`). pptxjs is a renderer — it reads a fixture
rather than authoring one — but it must agree with `python-pptx` on what
the feature's XML looks like.

1. **Read the manifest.** Look under
   `../ooxml-reference-corpus/features/pptx/` for a JSON manifest
   covering the feature you're rendering.

2. **Consult the ECMA-376 5th edition spec** (corpus-only):
   - PDFs: `../ooxml-reference-corpus/spec/ecma-376-5/part-{1,2,3,4}/*.pdf`
   - RNC schemas (easier to read): `../ooxml-reference-corpus/spec/ecma-376-5/part-1/rnc/`
   - XSD schemas: `../ooxml-reference-corpus/spec/ecma-376-5/part-1/xsd/`

3. **If no manifest exists**, ask python-pptx's maintainer to author one
   first — authoring-side libraries own the definition of "done".

4. **Verify rendering.** Add a Playwright render-smoke test loading the
   committed fixture from the corpus.

## Workflow

- Feature branches: `feat/<short-name>`. Test/tooling: `test/<short-name>`.
- Open one PR per logical change.

### Keep README.md and TODO.md current

Whenever a feature is added, removed, or a public option changes, update both of these files *in the same PR* as the code change — stale docs have bitten us before.

- **`README.md`** — the API block reflects the real public surface. If you add/remove a function or option, add/remove the matching entry. If you add or remove an export, reflect it in the API section. Any prose sections (Status, Contributing, project-specific sections) should also match reality.
- **`TODO.md`** — if the change resolves a tracked issue, move that entry into a "Resolved in fork" / "Done" section with a one-line description and the PR/commit reference. Update any counts table at the top and bump the "last updated" date.

Minimum check before every PR that touches source: `grep -n "<feature name>" README.md TODO.md` to catch stale references.
