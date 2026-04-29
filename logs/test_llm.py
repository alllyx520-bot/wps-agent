# -*- coding: utf-8 -*-
"""Phase 2 LLM-powered features verification."""
import asyncio, json, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from mcp_server import call_tool

print("="*60)
print("PHASE 2 LLM VERIFICATION")
print("="*60)

# Create a structured test document
print("\n[1] Create and structure test document")
r = asyncio.run(call_tool("document", {"action": "create"}))
j = json.loads(r[0].text)
print(f"  Created: {j.get('name','')}")

r = asyncio.run(call_tool("content", {"action": "insert_text", "text": (
    "第一章 项目概述\n"
    "本项目旨在建立一个智能办公平台，利用人工智能技术提升文档处理效率。\n"
    "项目背景源于企业对高效文档管理的迫切需求，当前人工排版耗时费力，\n"
    "急需自动化解决方案来解放生产力。\n\n"
    "1.1 技术选型\n"
    "系统采用Python作为主要开发语言，利用COM接口与WPS Office深度集成。\n"
    "前端采用WPS加载项技术，后端为MCP协议服务。\n\n"
    "1.2 系统架构\n"
    "系统分为三层架构：展示层、业务逻辑层和数据访问层。\n"
    "展示层负责用户交互，业务逻辑层处理文档操作，数据层管理持久化。\n\n"
    "第二章 详细设计\n"
    "详细设计部分将阐述各个模块的具体实现方案。\n\n"
    "2.1 文档解析模块\n"
    "文档解析模块负责读取WPS文档的结构化信息，包括段落、字体、样式等。\n\n"
    "2.2 格式转换模块\n"
    "格式转换模块将用户的自然语言指令转换为具体的COM操作序列。\n\n"
    "第三章 实施计划\n"
    "本章介绍项目的实施步骤和时间安排。\n"
)}))
j = json.loads(r[0].text)
print(f"  Content: {j.get('inserted')}")

# Set outline levels
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.GetObject(None, "Kwps.Application")
doc = app.ActiveDocument
doc.Paragraphs.Item(1).Format.OutlineLevel = 1
doc.Paragraphs.Item(5).Format.OutlineLevel = 2
doc.Paragraphs.Item(8).Format.OutlineLevel = 2
doc.Paragraphs.Item(11).Format.OutlineLevel = 1
doc.Paragraphs.Item(13).Format.OutlineLevel = 2
doc.Paragraphs.Item(15).Format.OutlineLevel = 2
doc.Paragraphs.Item(17).Format.OutlineLevel = 1
print("  Outline levels set: 3 chapters, 4 sections")

# Test LLM analyze
print("\n[2] ai_format.analyze (LLM-powered)")
r = asyncio.run(call_tool("ai_format", {"action": "analyze"}))
j = json.loads(r[0].text)
doc_info = j.get('document',{})
llm = j.get('llm_analysis',{})
print(f"  Paragraphs: {doc_info.get('paragraph_count')}")
print(f"  Outline: {len(j.get('outline',[]))} headings")
if llm and llm != "null":
    print(f"  LLM doc_type: {llm.get('doc_type','?')}")
    headings = llm.get('heading_hierarchy',[])
    for h in headings:
        print(f"    Level {h.get('level')}: {h.get('font','?')} x{h.get('count',0)}")
    incons = llm.get('inconsistencies',[])
    if incons:
        print(f"  Inconsistencies: {len(incons)} found")
        for inc in incons[:3]:
            print(f"    - {inc.get('description','')[:80]}")
    else:
        print(f"  No inconsistencies found")
else:
    print(f"  LLM: not available (check API key)")

# Test LLM suggest
print("\n[3] ai_format.suggest (LLM-powered)")
r = asyncio.run(call_tool("ai_format", {"action": "suggest"}))
j = json.loads(r[0].text)
suggestions = j.get('llm_suggestions',[])
if suggestions:
    print(f"  Suggestions: {len(suggestions)} actions")
    for s in suggestions[:5]:
        print(f"    {s.get('tool','?')}/{s.get('action','?')}: {s.get('reason','')[:60]}")
else:
    print(f"  LLM suggestions: not available")

# Test LLM reformat
print("\n[4] ai_format.reformat (LLM-powered)")
r = asyncio.run(call_tool("ai_format", {"action": "reformat", 
    "instructions": "将第一章和第二章的标题改为黑体三号加粗居中，正文全部改为宋体小四号两端对齐首行缩进两字符"}))
j = json.loads(r[0].text)
print(f"  Executed: {j.get('executed',0)}, Failed: {j.get('failed',0)}")
details = j.get('details',[])
for d in details[:5]:
    print(f"    {d[:80]}")

# Verify
print("\n[5] Verify changes")
r = asyncio.run(call_tool("format", {"action": "get_font", "para_index": 1}))
j = json.loads(r[0].text)
print(f"  Chapter 1: {j.get('name')} {j.get('size')}pt bold={j.get('bold')}")

r = asyncio.run(call_tool("format", {"action": "get_font", "para_index": 5}))
j = json.loads(r[0].text)
print(f"  Section 1.1: {j.get('name')} {j.get('size')}pt")

r = asyncio.run(call_tool("format", {"action": "get_font", "para_index": 2}))
j = json.loads(r[0].text)
print(f"  Body para: {j.get('name')} {j.get('size')}pt")

# Cleanup
r = asyncio.run(call_tool("document", {"action": "close", "save_changes": False}))

print("\n"+"="*60)
print("PHASE 2 LLM VERIFICATION COMPLETE")
print("="*60)
