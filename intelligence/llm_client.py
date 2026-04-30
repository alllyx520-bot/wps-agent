# -*- coding: utf-8 -*-
import os
import json
import logging
from pathlib import Path
from typing import Optional, Dict, List, AsyncIterator
import httpx

logger = logging.getLogger("wps-agent.llm")

try:
    import yaml
    _config_path = Path(__file__).parent.parent / "config.yaml"
    with open(_config_path, "r", encoding="utf-8") as f:
        _config = yaml.safe_load(f) or {}
except Exception:
    _config = {}

LLM_CONFIG = _config.get("llm", {})
ENDPOINT = LLM_CONFIG.get("endpoint", "https://dashscope.aliyuncs.com/compatible-mode/v1")
MODEL = LLM_CONFIG.get("model", "deepseek-v4-pro")
MAX_TOKENS = LLM_CONFIG.get("max_tokens", 4096)
TEMPERATURE = LLM_CONFIG.get("temperature", 0.3)


import re

def _strip_code_fence(text: str) -> str:
    m = re.search(r'```(?:\w+)?\s*\n?(.*?)\n?```', text, re.DOTALL)
    if m:
        return m.group(1).strip()
    return text.strip()


def _get_api_key() -> Optional[str]:
    key = LLM_CONFIG.get("api_key")
    if key:
        return key
    for env_name in ["WPS_AGENT_LLM_KEY", "DASHSCOPE_API_KEY", "ALIYUN_API_KEY", "OPENAI_API_KEY", "LLM_API_KEY", "API_KEY"]:
        key = os.environ.get(env_name)
        if key:
            return key
    return None


def _build_messages(system_prompt: str, user_prompt: str) -> List[Dict]:
    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]


def chat(system_prompt: str, user_prompt: str, model: Optional[str] = None,
         temperature: Optional[float] = None, max_tokens: Optional[int] = None) -> Optional[str]:
    api_key = _get_api_key()
    if not api_key:
        logger.error("No API key configured. Set DASHSCOPE_API_KEY env var or api_key in config.yaml")
        return None

    messages = _build_messages(system_prompt, user_prompt)
    url = f"{ENDPOINT.rstrip('/')}/chat/completions"

    try:
        with httpx.Client(timeout=120) as client:
            resp = client.post(url, json={
                "model": model or MODEL,
                "messages": messages,
                "temperature": temperature if temperature is not None else TEMPERATURE,
                "max_tokens": max_tokens or MAX_TOKENS,
            }, headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            })
            if resp.status_code != 200:
                logger.error(f"LLM API error {resp.status_code}: {resp.text[:200]}")
                return None
            data = resp.json()
            return data["choices"][0]["message"]["content"]
    except Exception as e:
        logger.error(f"LLM API call failed: {e}")
        return None


async def chat_async(system_prompt: str, user_prompt: str, model: Optional[str] = None,
                     temperature: Optional[float] = None, max_tokens: Optional[int] = None) -> Optional[str]:
    api_key = _get_api_key()
    if not api_key:
        logger.error("No API key configured")
        return None

    messages = _build_messages(system_prompt, user_prompt)
    url = f"{ENDPOINT.rstrip('/')}/chat/completions"

    try:
        async with httpx.AsyncClient(timeout=120) as client:
            resp = await client.post(url, json={
                "model": model or MODEL,
                "messages": messages,
                "temperature": temperature if temperature is not None else TEMPERATURE,
                "max_tokens": max_tokens or MAX_TOKENS,
            }, headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            })
            if resp.status_code != 200:
                logger.error(f"LLM API error {resp.status_code}: {resp.text[:200]}")
                return None
            data = resp.json()
            return data["choices"][0]["message"]["content"]
    except Exception as e:
        logger.error(f"LLM API call failed: {e}")
        return None


def analyze_document_structure(outline: List[Dict], format_samples: List[Dict]) -> Optional[Dict]:
    sys_prompt = """You are a professional document formatting expert for Chinese documents.
Analyze the document structure and output a JSON object with:
{
  "doc_type": "official"|"thesis"|"report"|"general",
  "heading_hierarchy": [{"level": 1, "font": "...", "count": N}, ...],
  "inconsistencies": [{"description": "...", "paragraphs": [N, ...]}, ...],
  "suggestions": [{"action": "...", "description": "...", "priority": "high"|"medium"|"low"}, ...]
}
Only output valid JSON, no extra text."""

    user_prompt = f"""Document Outline:
{json.dumps(outline, ensure_ascii=False, indent=2)}

Format Samples:
{json.dumps(format_samples, ensure_ascii=False, indent=2)}

Analyze the document structure and formatting, identify inconsistencies, and suggest improvements."""

    result = chat(sys_prompt, user_prompt)
    if not result:
        return None
    try:
        result = _strip_code_fence(result)
        return json.loads(result)
    except json.JSONDecodeError:
        return {"raw_analysis": result}


def suggest_formatting(outline: List[Dict], format_samples: List[Dict]) -> Optional[Dict]:
    sys_prompt = """You are a Chinese document formatting expert. 
Given a document's outline and format samples, suggest specific formatting actions.
Output a JSON array of formatting actions:
[
  {"tool": "format", "action": "set_font", "para_index": N, "name": "黑体", "size": 22, "bold": true, "reason": "..."},
  ...
]
IMPORTANT: Chinese documents NEVER use italic (italic) for headings or body text. Use bold for emphasis, not italic. Do NOT include "italic": true in any output. Only output valid JSON array, no extra text."""

    user_prompt = f"""Outline:
{json.dumps(outline, ensure_ascii=False, indent=2)}

Current Format Samples:
{json.dumps(format_samples, ensure_ascii=False, indent=2)}

Generate specific formatting actions to improve this document to professional standards."""

    result = chat(sys_prompt, user_prompt)
    if not result:
        return None
    try:
        result = _strip_code_fence(result)
        return json.loads(result)
    except json.JSONDecodeError:
        return {"raw_suggestions": result}


def parse_natural_language_instructions(instructions: str, outline: List[Dict],
                                        format_samples: List[Dict]) -> Optional[List[Dict]]:
    sys_prompt = """You are an expert at converting natural language formatting instructions into specific WPS/Word COM operations.
Give the user's instructions and document context, output a JSON array of tool calls.

IMPORTANT: Combine operations efficiently. Use format.batch for multiple paragraphs instead of individual calls.
CRITICAL: Chinese documents NEVER use italic (italic / Italic) for headings or body text. Use bold (bold / Bold) for emphasis. Do NOT include "italic": true or "Italic": true in any output.

Font names: 黑体, 宋体, 仿宋, 楷体, 微软雅黑, 方正小标宋简体, 等线, Calibri, Arial, Times New Roman
Alignments: left, center, right, justify
Line spacing rules: single, 1.5lines, double, exactly, multiple
Style names: 正文, 标题 1, 标题 2, 标题 3, 标题 (Title)

FEW-SHOT EXAMPLES:

Example 1:
Instruction: "把第一段改成黑体三号加粗居中"
Output: [{"tool": "format", "action": "set_font", "para_index": 1, "name": "黑体", "size": 16, "bold": true}, {"tool": "format", "action": "set_paragraph_format", "para_index": 1, "alignment": "center"}]

Example 2:
Instruction: "把所有标题改成黑体"
Output: [{"tool": "format", "action": "batch", "operations": [{"type": "set_font", "para_index": N, "name": "黑体"}]}]

Example 3:
Instruction: "设置A4纸张页边距上下2.54cm左右3.17cm"
Output: [{"tool": "layout", "action": "page_setup", "top_margin": 72, "bottom_margin": 72, "left_margin": 90, "right_margin": 90}]

Example 4:
Instruction: "正文用宋体小四号首行缩进两个字符"
Output: [{"tool": "format", "action": "apply_style", "style_name": "正文"}, {"tool": "format", "action": "set_font", "name": "宋体", "size": 12}, {"tool": "format", "action": "set_paragraph_format", "first_line_indent": 28}]

Only output valid JSON array, no extra text."""

    user_prompt = f"""User Instruction: {instructions}

Document Context:
Outline: {json.dumps(outline[:20], ensure_ascii=False)}
Format Samples: {json.dumps(format_samples[:10], ensure_ascii=False)}

Convert the user's natural language instruction into specific formatting operations."""

    result = chat(sys_prompt, user_prompt)
    if not result:
        return None
    try:
        result = _strip_code_fence(result)
        return json.loads(result)
    except json.JSONDecodeError:
        return None
