# pptxjs — project notes for Claude

Browser-side PPTX → HTML renderer. TypeScript, built with rollup, tested with Karma + jasmine. Follows the same structure as the sibling [docxjs](https://github.com/loadfix/docxjs) project.

## Workflow

- Feature branches: `feat/<short-name>`. Test/tooling: `test/<short-name>`.
- Open one PR per logical change.

### Keep README.md and TODO.md current

Whenever a feature is added, removed, or a public option changes, update both of these files *in the same PR* as the code change — stale docs have bitten us before.

- **`README.md`** — the API block reflects the real public surface. If you add/remove a function or option, add/remove the matching entry. If you add or remove an export, reflect it in the API section. Any prose sections (Status, Contributing, project-specific sections) should also match reality.
- **`TODO.md`** — if the change resolves a tracked issue, move that entry into a "Resolved in fork" / "Done" section with a one-line description and the PR/commit reference. Update any counts table at the top and bump the "last updated" date.

Minimum check before every PR that touches source: `grep -n "<feature name>" README.md TODO.md` to catch stale references.
