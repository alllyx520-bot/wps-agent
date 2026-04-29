# -*- coding: utf-8 -*-
"""End-to-end verification of all Phase 1 tools via direct COM + MCP simulation."""
import asyncio, json, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
from mcp_server import list_tools, call_tool

pythoncom.CoInitialize()

print("=" * 60)
print("WPS AGENT PHASE 1 - END-TO-END VERIFICATION")
print("=" * 60)

# --- Test 1: Tool listing ---
print("\n[1] Tool Count")
tools = asyncio.run(list_tools())
print(f"  Tools registered: {len(tools)}")
tool_names = [t.name for t in tools]
expected = ["document", "content", "format", "style", "table", "search", "layout", "review", "ai_format"]
for e in expected:
    if e in tool_names:
        print(f"    {e}: present")
    else:
        print(f"    {e}: MISSING!")

# --- Test 2: Document tools ---
print("\n[2] Document Tools")
r = asyncio.run(call_tool("document", {"action": "create"}))
d = json.loads(r[0].text)
print(f"  create: {d.get('name', d)}")

r = asyncio.run(call_tool("document", {"action": "info"}))
d = json.loads(r[0].text)
print(f"  info: paragraphs={d.get('paragraph_count')}, tables={d.get('table_count')}")

r = asyncio.run(call_tool("document", {"action": "list"}))
d = json.loads(r[0].text)
print(f"  list: {len(d)} document(s)")

# --- Test 3: Content tools ---
print("\n[3] Content Tools")
# Add text first via COM
import win32com.client
app = win32com.client.GetObject(None, "Kwps.Application")
doc = app.ActiveDocument
doc.Content.Text = "第一章 测试标题\n这是第一段正文内容，测试字体格式。\n这是第二段正文内容，测试段落格式。\n\n第二章 第二个标题\n这是第三段正文。\n"
doc.Paragraphs.Item(1).Format.OutlineLevel = 1
doc.Paragraphs.Item(5).Format.OutlineLevel = 1

r = asyncio.run(call_tool("content", {"action": "paragraph", "para_index": 1}))
d = json.loads(r[0].text)
print(f"  paragraph[1]: text='{d.get('text','')[:40]}', font={d.get('font',{}).get('name')}")

r = asyncio.run(call_tool("content", {"action": "outline"}))
d = json.loads(r[0].text)
print(f"  outline: {len(d)} headings")

r = asyncio.run(call_tool("content", {"action": "full_text"}))
d = json.loads(r[0].text)
print(f"  full_text: {len(d.get('text',''))} chars")

# --- Test 4: Format tools ---
print("\n[4] Format Tools")
r = asyncio.run(call_tool("format", {"action": "set_font", "para_index": 1, "name": "黑体", "size": 22, "bold": True}))
d = json.loads(r[0].text)
print(f"  set_font: {d.get('updated')}, text='{d.get('text_sample','')[:30]}'")

r = asyncio.run(call_tool("format", {"action": "get_font", "para_index": 1}))
d = json.loads(r[0].text)
fok = d.get('name') == '黑体' and d.get('size') == 22.0
print(f"  get_font: name={d.get('name')}, size={d.get('size')}, bold={d.get('bold')} {'OK' if fok else 'FAIL'}")

r = asyncio.run(call_tool("format", {"action": "set_paragraph_format", "para_index": 2, "alignment": "justify", "first_line_indent": 28, "line_spacing_rule": "multiple", "line_spacing": 1.5}))
d = json.loads(r[0].text)
print(f"  set_paragraph_format: updated={d.get('updated')}")

r = asyncio.run(call_tool("format", {"action": "get_paragraph_format", "para_index": 2}))
d = json.loads(r[0].text)
pfok = d.get('alignment') == 'justify' and d.get('line_spacing') == 1.5
print(f"  get_paragraph_format: align={d.get('alignment')}, indent={d.get('first_line_indent')}, spacing={d.get('line_spacing')} {'OK' if pfok else 'FAIL'}")

# --- Test 5: Style tools ---
print("\n[5] Style Tools")
r = asyncio.run(call_tool("style", {"action": "get", "name": "标题 1"}))
d = json.loads(r[0].text)
print(f"  get '标题 1': type={d.get('type')}, font={d.get('font',{}).get('name')}")

r = asyncio.run(call_tool("style", {"action": "create", "name": "MyTestStyle", "font_name": "微软雅黑", "font_size": 14, "bold": True}))
d = json.loads(r[0].text)
print(f"  create: {d.get('created', d)}")

# --- Test 6: Table tools ---
print("\n[6] Table Tools")
r = asyncio.run(call_tool("table", {"action": "create", "rows": 3, "cols": 3, "position": "end"}))
d = json.loads(r[0].text)
tidx = d.get('table_index')
print(f"  create: index={tidx}, {d.get('rows')}x{d.get('columns')}")

r = asyncio.run(call_tool("table", {"action": "set_cell_text", "table_index": tidx, "row": 1, "col": 1, "text": "Name"}))
d = json.loads(r[0].text)
print(f"  set_cell_text: row={d.get('row')}, col={d.get('col')}")

r = asyncio.run(call_tool("table", {"action": "read", "table_index": tidx}))
d = json.loads(r[0].text)
print(f"  read: {d.get('rows')}x{d.get('columns')}, cell[1,1]='{d.get('data',[[]])[0][0]}'")

# --- Test 7: Search tools ---
print("\n[7] Search Tools")
r = asyncio.run(call_tool("search", {"action": "find", "query": "测试"}))
d = json.loads(r[0].text)
print(f"  find '测试': {len(d)} matches")

r = asyncio.run(call_tool("search", {"action": "replace", "find_text": "测试", "replace_text": "验证", "replace_all": True}))
d = json.loads(r[0].text)
print(f"  replace '测试'→'验证': {d}")

# --- Test 8: Layout tools ---
print("\n[8] Layout Tools")
r = asyncio.run(call_tool("layout", {"action": "section_info"}))
d = json.loads(r[0].text)
print(f"  section_info: {d.get('page_width')}x{d.get('page_height')}, margins={d.get('top_margin')}/{d.get('bottom_margin')}/{d.get('left_margin')}/{d.get('right_margin')}")

r = asyncio.run(call_tool("layout", {"action": "page_setup", "top_margin": 36, "bottom_margin": 36, "left_margin": 28, "right_margin": 28}))
d = json.loads(r[0].text)
print(f"  page_setup: margins updated, failed={d.get('failed')}")

# --- Test 9: Review tools ---
print("\n[9] Review Tools")
r = asyncio.run(call_tool("review", {"action": "track_changes_toggle", "enable": True}))
d = json.loads(r[0].text)
print(f"  track_changes: {d}")

r = asyncio.run(call_tool("review", {"action": "comment_add", "text": "这是一条测试批注", "para_index": 1}))
d = json.loads(r[0].text)
print(f"  comment_add: {d.get('comment_added','')}")

r = asyncio.run(call_tool("review", {"action": "comments_list"}))
d = json.loads(r[0].text)
print(f"  comments_list: {len(d)} comment(s)")

# --- Test 10: AI Format template ---
print("\n[10] AI Format Tools")
r = asyncio.run(call_tool("ai_format", {"action": "analyze"}))
d = json.loads(r[0].text)
print(f"  analyze: paragraphs={d.get('document',{}).get('paragraph_count')}, outline={len(d.get('outline',[]))}")

r = asyncio.run(call_tool("ai_format", {"action": "apply_template", "template_name": "thesis"}))
d = json.loads(r[0].text)
print(f"  apply_template(thesis): applied={d.get('applied_to', d)}")

# --- Test 11: Save test ---
print("\n[11] Save")
test_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs", "test_e2e_output.docx")
r = asyncio.run(call_tool("document", {"action": "save", "filepath": test_path}))
d = json.loads(r[0].text)
print(f"  save: {d.get('full_name', d)}")

# Cleanup
doc.Close(False)
app.Quit()

print("\n" + "=" * 60)
print("PHASE 1 END-TO-END VERIFICATION COMPLETE")
print("=" * 60)
pythoncom.CoUninitialize()
