# -*- coding: utf-8 -*-
"""Phase 2 verification test."""
import asyncio, json, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from mcp_server import call_tool

print("="*60)
print("PHASE 2 VERIFICATION")
print("="*60)

# Create fresh doc with structured content
print("\n[1] Create test document")
r = asyncio.run(call_tool("document", {"action": "create"}))
j = json.loads(r[0].text)
print(f"  Created: {j.get('name','')}")

r = asyncio.run(call_tool("content", {"action": "insert_text", "text": "第一章 项目概述\n这是项目概述的正文内容，描述了项目的背景和目标。\n\n1.1 技术方案\n本节介绍技术方案的具体细节。\n\n1.2 实施计划\n本节介绍实施计划和时间安排。\n\n第二章 详细设计\n这是详细设计的正文内容。\n\n2.1 架构设计\n介绍系统架构。\n\n2.2 模块设计\n介绍各模块设计。\n\n第三章 总结\n这是总结内容。\n"}))
j = json.loads(r[0].text)
print(f"  Content inserted: {j.get('inserted')}")

# Set outline levels (MCP uses its own COM, don't CoUninitialize)
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.GetObject(None, "Kwps.Application")
doc = app.ActiveDocument
doc.Paragraphs.Item(1).Format.OutlineLevel = 1
doc.Paragraphs.Item(3).Format.OutlineLevel = 2
doc.Paragraphs.Item(5).Format.OutlineLevel = 2
doc.Paragraphs.Item(8).Format.OutlineLevel = 1
doc.Paragraphs.Item(10).Format.OutlineLevel = 2
doc.Paragraphs.Item(12).Format.OutlineLevel = 2
doc.Paragraphs.Item(15).Format.OutlineLevel = 1
# Do NOT call CoUninitialize - MCP server needs COM alive
print("  Outline levels set")

# Test analyze
print("\n[2] ai_format.analyze")
r = asyncio.run(call_tool("ai_format", {"action": "analyze"}))
j = json.loads(r[0].text)
print(f"  Document: {j.get('document',{}).get('paragraph_count')} paragraphs")
print(f"  Outline: {len(j.get('outline',[]))} headings")
samples = j.get('format_samples', [])
for s in samples:
    print(f"    [{s['outline_level']}] {s['text'][:40]} | font={s['font']['name']} {s['font']['size']}pt")

# Test auto_numbering
print("\n[3] ai_format.auto_numbering")
r = asyncio.run(call_tool("ai_format", {"action": "auto_numbering"}))
j = json.loads(r[0].text)
print(f"  Numbered: {j.get('numbered_headings', 0)} headings")

# Verify numbering
r = asyncio.run(call_tool("content", {"action": "paragraph", "para_index": 3}))
j = json.loads(r[0].text)
print(f"  Para[3] after numbering: '{j.get('text','')[:60]}'")

# Test apply_template
print("\n[4] ai_format.apply_template (thesis)")
r = asyncio.run(call_tool("ai_format", {"action": "apply_template", "template_name": "thesis"}))
j = json.loads(r[0].text)
print(f"  Applied: {j.get('applied_to', j)}")

# Verify formatting
r = asyncio.run(call_tool("format", {"action": "get_font", "para_index": 1}))
j = json.loads(r[0].text)
print(f"  Chapter heading font: {j.get('name')} {j.get('size')}pt bold={j.get('bold')}")

r = asyncio.run(call_tool("format", {"action": "get_font", "para_index": 2}))
j = json.loads(r[0].text)
print(f"  Body font: {j.get('name')} {j.get('size')}pt")

# Test auto_toc
print("\n[5] ai_format.auto_toc")
r = asyncio.run(call_tool("ai_format", {"action": "auto_toc"}))
j = json.loads(r[0].text)
print(f"  TOC: {j}")

# Test validate
print("\n[6] ai_format.validate")
r = asyncio.run(call_tool("ai_format", {"action": "validate"}))
j = json.loads(r[0].text)
print(f"  Issues: {j.get('issues_found', 0)}, outline: {j.get('outline_count', 0)} headings")

# Test suggest
print("\n[7] ai_format.suggest")
r = asyncio.run(call_tool("ai_format", {"action": "suggest"}))
j = json.loads(r[0].text)
print(f"  Samples: {len(j.get('format_samples',[]))}")
if j.get('llm_suggestions'):
    print(f"  LLM suggestions: {j['llm_suggestions']}")
else:
    print(f"  LLM: API key not configured (expected)")

# Test reformat (no LLM)
print("\n[8] ai_format.reformat (no LLM)")
r = asyncio.run(call_tool("ai_format", {"action": "reformat", "instructions": "将所有标题改为黑体三号加粗居中，正文改为宋体小四号"}))
j = json.loads(r[0].text)
print(f"  Executed: {j.get('executed',0)}, Failed: {j.get('failed',0)}")
if j.get('note'):
    print(f"  Note: {j['note']}")

# Cleanup via MCP
r = asyncio.run(call_tool("document", {"action": "close", "save_changes": False}))

print("\n"+"="*60)
print("PHASE 2 VERIFICATION COMPLETE")
print("="*60)
