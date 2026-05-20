#!/usr/bin/env python3
from __future__ import annotations

import json
import shlex
import sys
from datetime import UTC, datetime
from pathlib import Path

PIPELINE_SCRIPT = "thesis_format_pipeline.py"


def find_root(cwd: str) -> Path:
    path = Path(cwd).resolve()
    for candidate in [path, *path.parents]:
        if (candidate / ".codex").exists() or (candidate / ".git").exists():
            return candidate
    return path


def load_registry(root: Path) -> dict:
    registry_path = root / "artifacts" / ".workflow" / "thesis-format-registry.json"
    if not registry_path.exists():
        return {}
    return json.loads(registry_path.read_text(encoding="utf-8"))


def approval_path(root: Path) -> Path:
    return root / "artifacts" / ".workflow" / "repair-approval.json"


def load_approval(root: Path) -> dict | None:
    path = approval_path(root)
    if not path.exists():
        return None
    approval = json.loads(path.read_text(encoding="utf-8"))
    if not approval.get("active"):
        return None
    return approval


def save_approval(root: Path, approval: dict) -> None:
    path = approval_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(approval, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def emit(payload: dict) -> int:
    print(json.dumps(payload, ensure_ascii=False))
    return 0


def deny(reason: str) -> int:
    return emit(
        {
            "hookSpecificOutput": {
                "hookEventName": "PreToolUse",
                "permissionDecision": "deny",
                "permissionDecisionReason": reason,
            }
        }
    )


def add_context(message: str) -> int:
    return emit(
        {
            "hookSpecificOutput": {
                "hookEventName": "PreToolUse",
                "additionalContext": message,
            }
        }
    )


def parse_command(command: str) -> list[str]:
    try:
        return shlex.split(command, posix=False)
    except ValueError:
        return command.split()


def pipeline_index(tokens: list[str]) -> int | None:
    for index, token in enumerate(tokens):
        if token.replace("\\", "/").endswith(PIPELINE_SCRIPT):
            return index
    return None


def flag_value(tokens: list[str], flag: str) -> str | None:
    for index, token in enumerate(tokens):
        if token == flag and index + 1 < len(tokens):
            return tokens[index + 1]
        if token.startswith(f"{flag}="):
            return token.split("=", 1)[1]
    return None


def resolve_path(raw: str | None, root: Path) -> Path | None:
    if raw is None:
        return None
    path = Path(raw)
    if not path.is_absolute():
        path = (root / path).resolve()
    return path


def latest_rule_pack(registry: dict) -> Path | None:
    latest = registry.get("latest_rule_pack")
    if not latest:
        return None
    path = Path(latest["path"])
    return path if path.exists() else None


def rule_pack_validation_error(rule_pack_path: Path) -> str | None:
    manifest_path = rule_pack_path / "manifest.json"
    if not manifest_path.exists():
        return f"Rule pack manifest not found: {manifest_path}"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    validation = manifest.get("validation") or {}
    if validation.get("pass") and validation.get("comparison") == "docx-package-metadata":
        return None
    if validation.get("pass") and validation.get("comparison") != "docx-package-metadata":
        return "Rule pack validation is stale. Rerun validate-rule-pack with DOCX package metadata diff enabled."
    report = validation.get("report_md") or validation.get("report_json")
    if report:
        return f"Rule pack validation failed. Inspect {report} and rerun validate-rule-pack before continuing."
    return "Rule pack validation has not passed. Run validate-rule-pack before continuing."


def matching_audit(registry: dict, document_path: Path, rule_pack_path: Path) -> bool:
    document_path = document_path.resolve()
    rule_pack_path = rule_pack_path.resolve()
    for audit in reversed(registry.get("audits", [])):
        if (
            Path(audit["document_path"]).resolve() == document_path
            and Path(audit["rule_pack_path"]).resolve() == rule_pack_path
        ):
            return True
    return False


def main() -> int:
    payload = json.load(sys.stdin)
    tool_input = payload.get("tool_input") or {}
    command = tool_input.get("command")
    if not isinstance(command, str):
        return 0

    tokens = parse_command(command)
    script_index = pipeline_index(tokens)
    if script_index is None or script_index + 1 >= len(tokens):
        return 0

    root = find_root(payload.get("cwd") or ".")
    registry = load_registry(root)
    subcommand = tokens[script_index + 1]

    if subcommand == "build-rule-pack":
        template_value = flag_value(tokens, "--template")
        template_path = resolve_path(template_value, root)
        if template_path is None:
            return deny("Pass --template <template.docx> before building a thesis rule pack.")
        if template_path.suffix.lower() != ".docx":
            return deny("The template input must be a .docx file.")
        if not template_path.exists():
            return deny(f"Template DOCX not found: {template_path}")
        return add_context(
            "When building the rule pack, extract both measurable DOCX formatting and explicit prose-stated formatting requirements from the template paragraphs into the persisted rule pack."
        )

    if subcommand == "audit":
        document_value = flag_value(tokens, "--document")
        document_path = resolve_path(document_value, root)
        if document_path is None:
            return deny("Pass --document <thesis.docx> before auditing.")
        if not document_path.exists():
            return deny(f"Thesis DOCX not found: {document_path}")
        rule_pack_value = flag_value(tokens, "--rule-pack")
        rule_pack_path = resolve_path(rule_pack_value, root) if rule_pack_value else latest_rule_pack(registry)
        if rule_pack_path is None:
            return deny("Build a rule pack first or pass --rule-pack <dir> before auditing.")
        if not rule_pack_path.exists():
            return deny(f"Rule pack not found: {rule_pack_path}")
        validation_error = rule_pack_validation_error(rule_pack_path)
        if validation_error:
            return deny(validation_error)
        if rule_pack_value is None:
            return add_context(f"Using the latest registered rule pack: {rule_pack_path}")
        return 0

    if subcommand == "repair":
        document_value = flag_value(tokens, "--document")
        document_path = resolve_path(document_value, root)
        if document_path is None:
            return deny("Pass --document <thesis.docx> before repairing.")
        if not document_path.exists():
            return deny(f"Thesis DOCX not found: {document_path}")
        rule_pack_value = flag_value(tokens, "--rule-pack")
        rule_pack_path = resolve_path(rule_pack_value, root) if rule_pack_value else latest_rule_pack(registry)
        if rule_pack_path is None:
            return deny("Build a rule pack first or pass --rule-pack <dir> before repairing.")
        if not rule_pack_path.exists():
            return deny(f"Rule pack not found: {rule_pack_path}")
        validation_error = rule_pack_validation_error(rule_pack_path)
        if validation_error:
            return deny(validation_error)
        if not matching_audit(registry, document_path, rule_pack_path):
            return deny("Run the audit step with the same thesis DOCX and rule pack before repair.")

        approval_id = flag_value(tokens, "--approval-id")
        approval = load_approval(root)
        if approval is None:
            return deny("Before repair, ask the user whether to proceed and wait for an explicit approval turn.")
        if approval.get("used_at"):
            return deny("The stored repair approval has already been used. Ask the user again before running repair.")
        if not approval_id:
            return deny("Repair requires an --approval-id token from a user turn that explicitly approved repair.")
        if approval_id != approval.get("approval_id"):
            return deny("The provided --approval-id token is invalid for the current repair approval.")

        approved_document = Path(approval["document_path"]).resolve()
        approved_rule_pack = Path(approval["rule_pack_path"]).resolve()
        if document_path.resolve() != approved_document or rule_pack_path.resolve() != approved_rule_pack:
            return deny("The repair approval only applies to the latest audited thesis document and rule pack.")

        output_value = flag_value(tokens, "--output")
        output_path = resolve_path(output_value, root)
        if output_path is not None and output_path.resolve() == document_path.resolve():
            return deny("Repair output must not overwrite the input thesis DOCX in place.")

        approval["used_at"] = datetime.now(UTC).isoformat()
        approval["active"] = False
        save_approval(root, approval)
        messages = []
        if rule_pack_value is None:
            messages.append(f"Using the latest registered rule pack: {rule_pack_path}")
        messages.append("Validated one-time user approval for repair.")
        return add_context(" ".join(messages))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
