# -*- coding: utf-8 -*-
"""Operation Logger — records all MCP tool calls with timing and results for debugging and audit."""
import json
import logging
import time
from datetime import datetime
from typing import Dict, List, Optional, Any
from pathlib import Path

logger = logging.getLogger("wps-agent.ops-log")

_log: List[Dict] = []
_enabled: bool = True
_max_entries: int = 500


def enable():
    global _enabled
    _enabled = True


def disable():
    global _enabled
    _enabled = False


def record(tool: str, action: str, params: Dict, result: Dict, duration: float, error: Optional[str] = None):
    if not _enabled:
        return
    entry = {
        "timestamp": datetime.now().isoformat(),
        "tool": tool,
        "action": action,
        "params": _sanitize_params(params),
        "success": "error" not in str(result) if isinstance(result, dict) else True,
        "duration_ms": round(duration * 1000, 2),
        "error": error,
    }
    _log.append(entry)
    if len(_log) > _max_entries:
        _log.pop(0)
    logger.debug(f"[{tool}.{action}] {entry['duration_ms']}ms {'OK' if entry['success'] else 'FAIL'}")


def _sanitize_params(params: Dict) -> Dict:
    sanitized = {}
    for k, v in params.items():
        if isinstance(v, str) and len(v) > 200:
            sanitized[k] = v[:200] + "..."
        elif isinstance(v, dict):
            sanitized[k] = _sanitize_params(v)
        elif isinstance(v, list) and len(v) > 20:
            sanitized[k] = v[:20]
        else:
            sanitized[k] = v
    return sanitized


def get_recent(count: int = 20) -> List[Dict]:
    return _log[-count:]


def get_summary() -> Dict:
    if not _log:
        return {"total_calls": 0}
    tools = {}
    successes = 0
    failures = 0
    total_time = 0
    for entry in _log:
        name = entry["tool"]
        if name not in tools:
            tools[name] = {"calls": 0, "actions": {}, "total_ms": 0}
        tools[name]["calls"] += 1
        tools[name]["total_ms"] += entry["duration_ms"]
        action = entry["action"]
        tools[name]["actions"][action] = tools[name]["actions"].get(action, 0) + 1
        if entry["success"]:
            successes += 1
        else:
            failures += 1
        total_time += entry["duration_ms"]
    return {
        "total_calls": len(_log),
        "successes": successes,
        "failures": failures,
        "total_time_ms": round(total_time, 1),
        "by_tool": tools,
    }


def get_errors() -> List[Dict]:
    return [e for e in _log if not e["success"]]


def replay_last_error() -> Optional[Dict]:
    for e in reversed(_log):
        if not e["success"]:
            return e
    return None


def clear():
    _log.clear()


def dump(filepath: str) -> str:
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump({"summary": get_summary(), "entries": _log}, f, ensure_ascii=False, indent=2, default=str)
    return filepath


def load(filepath: str):
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    global _log
    _log = data.get("entries", [])
    return len(_log)
