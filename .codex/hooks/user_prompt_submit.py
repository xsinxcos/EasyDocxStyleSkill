#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sys
import uuid
from datetime import UTC, datetime
from pathlib import Path

KEYWORDS = (
    "thesis",
    "dissertation",
    "paper format",
    "论文",
    "毕业论文",
    "学位论文",
    "格式",
    "rule pack",
    "规则包",
)
APPROVAL_PATTERNS = (
    r"\brepair\b",
    r"\bgo ahead\b",
    r"\bproceed with repair\b",
    r"\brepair it\b",
    r"\bfix it\b",
    r"\bapply the repair\b",
    r"\bapply the fix\b",
    r"开始修复",
    r"执行修复",
    r"可以修复",
    r"请修复",
    r"继续修复",
    r"修复吧",
    r"修吧",
    r"可以，?修复",
    r"开始调整格式",
)
NEGATIVE_PATTERNS = (
    r"\bdo not repair\b",
    r"\bdon't repair\b",
    r"\bno repair\b",
    r"不要修复",
    r"先别修复",
    r"暂不修复",
)


def find_root(cwd: str) -> Path:
    path = Path(cwd).resolve()
    for candidate in [path, *path.parents]:
        if (candidate / ".codex").exists() or (candidate / ".git").exists():
            return candidate
    return path


def load_registry(root: Path) -> dict:
    registry = root / "artifacts" / ".workflow" / "thesis-format-registry.json"
    if not registry.exists():
        return {}
    return json.loads(registry.read_text(encoding="utf-8"))


def approval_path(root: Path) -> Path:
    return root / "artifacts" / ".workflow" / "repair-approval.json"


def clear_approval(root: Path) -> None:
    path = approval_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "active": False,
        "approval_id": None,
        "granted_at": None,
        "document_path": None,
        "rule_pack_path": None,
        "report_json": None,
        "used_at": None,
        "revoked_at": datetime.now(UTC).isoformat(),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def save_approval(root: Path, payload: dict) -> None:
    path = approval_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"active": True, **payload}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def is_thesis_prompt(prompt: str) -> bool:
    prompt = prompt.lower()
    return any(keyword in prompt for keyword in KEYWORDS)


def has_explicit_repair_approval(prompt: str) -> bool:
    lowered = prompt.lower()
    if any(re.search(pattern, lowered, re.IGNORECASE) for pattern in NEGATIVE_PATTERNS):
        return False
    return any(re.search(pattern, lowered, re.IGNORECASE) for pattern in APPROVAL_PATTERNS)


def main() -> int:
    payload = json.load(sys.stdin)
    prompt = payload.get("prompt") or ""
    root = find_root(payload.get("cwd") or ".")
    registry = load_registry(root)
    clear_approval(root)
    latest_audit = registry.get("latest_audit")
    explicit_repair_approval = has_explicit_repair_approval(prompt)
    if not is_thesis_prompt(prompt) and not (latest_audit and explicit_repair_approval):
        return 0

    context = [
        "This workspace contains a template-driven thesis formatting workflow.",
        "Use $thesis-format-enforcer and the pipeline under .agents/skills/thesis-format-enforcer/scripts.",
    ]
    latest_rule_pack = registry.get("latest_rule_pack")
    if latest_rule_pack:
        context.append(
            f"Latest rule pack: {latest_rule_pack['path']}. Reuse it for audit or repair unless the user asks for a new template."
        )
    else:
        context.append(
            "No rule pack is registered yet. Before auditing or repairing a thesis DOCX, require a template DOCX and run the build-rule-pack step."
        )

    context.append(
        "A rule pack is only usable after validate-rule-pack (or build-rule-pack's built-in validation) successfully rebuilds the template and reports a pass with DOCX package metadata diff."
    )
    context.append(
        "When building a rule pack, capture both the template's measurable DOCX formatting and any explicit formatting requirements written in the template text itself."
    )
    context.append("Never skip the audit step before the repair step.")
    if latest_audit:
        context.append(
            "Before running repair, ask the user after the audit and wait for explicit approval in a later turn."
        )
        if explicit_repair_approval:
            approval_id = uuid.uuid4().hex
            approval_payload = {
                "approval_id": approval_id,
                "granted_at": datetime.now(UTC).isoformat(),
                "document_path": latest_audit["document_path"],
                "rule_pack_path": latest_audit["rule_pack_path"],
                "report_json": latest_audit["report_json"],
                "used_at": None,
            }
            save_approval(root, approval_payload)
            context.append(
                f"User explicitly approved repair for the latest audit. Use --approval-id {approval_id} exactly once with the repair command for document {latest_audit['document_path']}."
            )
    else:
        context.append("No audited thesis is active yet, so no repair approval can be issued.")

    print(
        json.dumps(
            {
                "hookSpecificOutput": {
                    "hookEventName": "UserPromptSubmit",
                    "additionalContext": " ".join(context),
                }
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
