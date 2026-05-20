#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sys


def main() -> int:
    payload = json.load(sys.stdin)
    if payload.get("stop_hook_active"):
        return 0
    message = payload.get("last_assistant_message") or ""
    if not re.search(r"(论文|thesis)", message, re.IGNORECASE):
        return 0
    if not re.search(r"(规则包|rule pack|audit|审查|repair|修复)", message, re.IGNORECASE):
        return 0
    if re.search(r"(artifacts[\\/]|\.docx|\.json|\.md)", message, re.IGNORECASE):
        return 0
    print(
        json.dumps(
            {
                "decision": "block",
                "reason": "Name the exact generated artifact paths for the thesis workflow, or explicitly state that no artifact was written.",
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
