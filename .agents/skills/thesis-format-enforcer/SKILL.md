---
name: thesis-format-enforcer
description: Derive a persistent thesis DOCX rule pack from a university or journal template, then audit and repair thesis manuscripts against that rule pack. Use when Codex needs to inspect 学位论文 or 毕业论文 formatting, compare a draft `.docx` with a template `.docx`, generate reusable formatting rules, produce strict audit reports, or normalize a thesis document before submission.
---

# Thesis Format Enforcer

## Overview

Build a reusable rule pack from a template thesis DOCX, then force the rule pack to prove itself by rebuilding the template from the persisted rule-pack payload and comparing the rebuilt DOCX against the original template at the OOXML package metadata level. Only treat the rule pack as usable when the package diff reports zero differences. The rule pack must capture both measurable DOCX formatting and any explicit formatting requirements written in the template text itself, because many thesis templates describe rules in prose instead of encoding every rule structurally. After that, reuse the validated rule pack to audit and repair thesis drafts. Prefer the bundled pipeline instead of ad hoc formatting edits so the workflow stays reproducible and the repo-local hooks can enforce the order.

## Workflow

1. Resolve the starting state.
- If no rule pack exists, require a template DOCX first.
- If a compatible rule pack already exists, reuse it instead of re-inferring rules from memory.
- Read `references/workflow.md` only when you need the exact artifact layout or CLI contract.

2. Build the rule pack.
- Run `python scripts/thesis_format_pipeline.py build-rule-pack --template <template.docx> [--name <pack-name>]`.
- The build step must also rebuild the template from `rules.json` and write validation artifacts under the rule-pack directory.
- The build step must also persist a DOCX package snapshot and rebuild the template from the rule-pack payload before validating it.
- Inspect `manifest.json`, `rules.json`, `text-requirements.md`, `package-manifest.json`, `template-index.md`, and `validation/validation-report.md` in the generated rule-pack directory.
- Treat `rules.json` as the authoritative persisted contract. The rebuild step must use the persisted rule pack, not the original template snapshot, as its source.
- The validation must compare DOCX package parts and metadata, not just a high-level extracted model.
- Confirm that prose-stated requirements from the template were persisted under `text_requirement_rules` and, when supported, activated into `semantic_rules`, `tracked_styles`, or `custom_paragraph_rules`.
- If the validation report is not a pass, the rule pack is not usable yet. Fix the persisted rule-pack payload and rerun `python scripts/thesis_format_pipeline.py validate-rule-pack --rule-pack <dir> [--template <template.docx>]`.

3. Audit the thesis draft.
- Run `python scripts/thesis_format_pipeline.py audit --document <thesis.docx> [--rule-pack <dir>]`.
- Audit only with a validation-passed rule pack.
- Treat `audit-report.json` as the source of truth and `audit-report.md` as the human-facing summary.
- Report errors before warnings. Cite exact paragraph indices, section labels, or style names.
- After the audit, stop and ask the user whether to proceed with repair. Do not launch `repair` in the same turn that produced the audit unless the user answers in a later turn.

4. Repair only after an audit exists for the same thesis and rule pack.
- Run `python scripts/thesis_format_pipeline.py repair --document <thesis.docx> [--rule-pack <dir>] [--output <fixed.docx>] --approval-id <token>`.
- The `--approval-id` token must come from a user turn that explicitly approved repair after the audit.
- Review the post-repair audit artifacts before claiming the document is clean.
- If findings remain, explain which items still need manual intervention, such as pagination, section breaks, or content that cannot be inferred safely from the template.

## Operating Rules

- Keep rule packs persistent on disk. Reuse them across audit and repair steps.
- Do not audit or repair from memory when a rule pack can be built or loaded.
- Do not treat a freshly extracted pack as valid until template reconstruction passes with zero DOCX package metadata differences.
- Do not rely on visible style extraction alone. Read the template paragraphs that state formatting requirements and persist those requirements into the rule pack.
- The template reconstruction step must be driven by the persisted rule pack contents. Do not silently consult the original template structure to fill missing gaps during rebuild.
- Never run the repair step without asking the user first and receiving an explicit yes in a later turn.
- Prefer style-level and semantic-section fixes over paragraph-by-paragraph manual patching.
- Do not promise perfect pagination. This workflow is strict on measurable style and structure, but some page-flow adjustments still require manual review in Word or WPS.
- If the input is not `.docx`, convert it first or ask the user for a `.docx` source.
- Hooks are repo-local, not skill-local. If this skill is copied into another repo, also copy `.codex/hooks.json` and `.codex/hooks/` if you need the same workflow guardrails.

## Resources

- `scripts/thesis_format_pipeline.py`: main CLI for `build-rule-pack`, `validate-rule-pack`, `audit`, `repair`, and `status`.
- `scripts/docx_model.py`: deterministic DOCX extractor used by the pipeline.
- `references/workflow.md`: artifact locations, registry semantics, and command examples.
- `references/rule-pack-schema.md`: persisted rule-pack file layout and which files are authoritative.
