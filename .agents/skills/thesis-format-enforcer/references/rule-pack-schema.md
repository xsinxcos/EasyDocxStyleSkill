# Rule Pack Schema

Each generated rule pack is a directory. The directory itself is the stable reference that later `audit` and `repair` commands consume.

## Files

- `manifest.json`
  - Generator metadata.
  - Rule-pack name.
  - Template path and template SHA-256.
  - Generation timestamp.
  - Validation status, rebuilt-template artifact path, and validation report paths.
- `rules.json`
  - Authoritative persisted rules.
  - Document defaults, section settings, required sections, semantic paragraph rules, cover rules, tracked style definitions, and the template blueprint used for rebuild validation.
- `text-requirements.md`
  - Human-readable summary of template prose requirements that were extracted from template paragraphs.
  - Shows which requirements became active machine-checkable rules and which are stored for manual review only.
- `package-manifest.json`
  - DOCX package snapshot manifest.
  - Records each package part path, raw SHA-256, and byte size.
- `package-snapshot/`
  - Raw package parts copied from the template DOCX.
  - Used to rebuild the template for metadata-level validation.
- `template-model.json`
  - Full extracted snapshot of the template document.
  - Useful for debugging or for manually enriching `rules.json`.
- `template-index.md`
  - Human-readable paragraph index with paragraph number, style, semantic label, and text excerpt.
- `validation/rebuilt-template.docx`
  - DOCX reconstructed only from the persisted rule-pack payload.
- `validation/validation-report.json`
  - Machine-readable DOCX package diff between the original template and the rebuilt template.
- `validation/validation-report.md`
  - Human-readable validation summary.

## Rule categories

- `document_defaults`
  - Default run and paragraph properties from the template.
- `section_settings`
  - Page size, orientation, margins, header distance, and footer distance.
- `required_sections`
  - Ordered semantic sections detected in the template.
- `semantic_rules`
  - First detected paragraph for semantic anchors such as `abstract_cn`, `toc`, or `references`.
- `cover_rules.blocks`
  - Front-matter block sequence before the first main semantic anchor.
- `tracked_styles`
  - Style definitions that the audit and repair steps should keep aligned with the template.
- `header_rules`
  - Header paragraph rules extracted from the template headers and optionally strengthened by prose requirements.
- `custom_table_rules`
  - Table-cell paragraph rules synthesized from prose requirements such as table body font and size.
- `template_blueprint`
  - Ordered paragraph and run sequence used to rebuild the template from the rule pack alone.
- `requirement_notes`
  - Short template paragraphs that look like explicit formatting instructions.
- `text_requirement_rules`
  - Parsed machine-readable requirements derived from `requirement_notes`.
  - Includes the source paragraph index, parsed run/paragraph spec, support status, and where the requirement should be applied.
- `custom_paragraph_rules`
  - Scoped paragraph/runs checks synthesized from prose requirements.
  - Used for checks such as abstract body formatting, body paragraph formatting, and reference entry formatting.

## Editing guidance

- Edit `rules.json` when the extractor misses a rule that is clearly required by the template.
- Prefer editing `text_requirement_rules` or the downstream rules they activate when the missing rule originates from prose inside the template.
- After editing `rules.json`, rerun `validate-rule-pack`; do not assume the pack is still valid.
- Keep `manifest.json` machine-generated.
- Use `template-model.json` and `template-index.md` only as supporting evidence; do not treat them as the final contract.
