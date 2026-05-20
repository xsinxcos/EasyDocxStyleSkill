# Thesis Workflow

## Artifact layout

The pipeline writes predictable artifacts under the workspace root:

- `artifacts/rule-packs/<pack-name>/`
- `artifacts/rule-packs/<pack-name>/validation/`
- `artifacts/rule-packs/<pack-name>/text-requirements.md`
- `artifacts/audits/<document-stem>/`
- `artifacts/repaired/<document-stem>-formatted.docx`
- `artifacts/repairs/<document-stem>/`
- `artifacts/.workflow/thesis-format-registry.json`

Treat `artifacts/.workflow/thesis-format-registry.json` as the workflow memory for the current repo. It records the latest rule pack, audits, and repair runs so the hooks can enforce the sequence.

## Commands

Build a reusable rule pack from the template thesis DOCX.

```bash
python scripts/thesis_format_pipeline.py build-rule-pack --template path/to/template.docx --name school-2026
```

The build command now extracts the rule pack, rebuilds the template from the persisted rules, compares the rebuilt DOCX against the original template, and writes validation artifacts under `validation/`. A rule pack is only usable after this validation passes.

The validation is package-level: compare DOCX part names plus each part's raw `sha256` and size. A zero-diff package report is the gate for a usable rule pack.

The build step also scans template paragraphs that explicitly describe formatting requirements. It persists the raw notes in `requirement_notes`, the parsed machine-readable rules in `text_requirement_rules`, and the activated scopes in `custom_paragraph_rules`, `custom_table_rules`, and `header_rules`. Inspect `text-requirements.md` to verify which prose requirements became executable checks and which remain manual-review items.

Rerun validation after editing `rules.json`.

```bash
python scripts/thesis_format_pipeline.py validate-rule-pack --rule-pack artifacts/rule-packs/school-2026 --template path/to/template.docx
```

Audit a thesis draft. If `--rule-pack` is omitted, the pipeline uses the latest registered rule pack.

```bash
python scripts/thesis_format_pipeline.py audit --document path/to/thesis.docx --rule-pack artifacts/rule-packs/school-2026
```

Repair a thesis draft only after an audit exists for the same draft and rule pack.

```bash
python scripts/thesis_format_pipeline.py repair --document path/to/thesis.docx --rule-pack artifacts/rule-packs/school-2026 --output artifacts/repaired/thesis-fixed.docx --approval-id <token>
```

The repair token is issued by the project hooks only after:

1. A latest audit exists in `artifacts/.workflow/thesis-format-registry.json`
2. Codex asks the user whether to proceed with repair
3. The user explicitly approves repair in a later turn

Inspect the registry state.

```bash
python scripts/thesis_format_pipeline.py status
```

## Hook contract

This repo ships project-local Codex hooks in `.codex/hooks.json`.

- `UserPromptSubmit` injects workflow context so Codex prefers the persisted rule pack flow and remembers that only validation-passed packs are usable.
- `PreToolUse` blocks out-of-order pipeline invocations, such as auditing without a rule pack, using an unvalidated or stale-validation rule pack, repairing before audit, or repairing without a fresh user approval token.
- `Stop` asks Codex to name the generated artifact paths before ending a turn when it reports thesis-format work.

Project-local hooks only load when the repo `.codex/` layer is trusted.

## Limits

- The extractor is strongest on measurable DOCX structure: styles, fonts, alignment, spacing, sections, and cover/front-matter block order.
- Prose-derived requirement parsing is heuristic. Unsupported template notes are still persisted in the rule pack, but they may remain `stored_only` until you add a deterministic mapping.
- The repair step does not guarantee perfect pagination, floating object placement, or every manual layout nuance in a university template.
- If the template contains prose requirements that are not expressed structurally, check `text-requirements.md`, enrich `rules.json` as needed, and rerun `validate-rule-pack`.
