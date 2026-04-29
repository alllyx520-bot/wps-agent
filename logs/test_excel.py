# -*- coding: utf-8 -*-
"""Excel COM integration verification."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from wps_bridge.excel_app import ExcelApplication

print("="*60)
print("EXCEL COM VERIFICATION")
print("="*60)

xl = ExcelApplication()
app = xl.app
print(f"Excel Version: {app.Version}")
print(f"Excel Name: {app.Name}")

# Create workbook
wb = app.Workbooks.Add()
ws = wb.Worksheets(1)
ws.Name = "测试数据"
print(f"Created: {wb.Name}, Sheet: {ws.Name}")

# Write data
data = [
    ["姓名", "部门", "销售额", "季度"],
    ["张三", "销售部", 120000, "Q1"],
    ["李四", "销售部", 95000, "Q1"],
    ["王五", "市场部", 80000, "Q1"],
    ["赵六", "技术部", 150000, "Q1"],
    ["张三", "销售部", 135000, "Q2"],
    ["李四", "销售部", 110000, "Q2"],
]
for i, row in enumerate(data):
    for j, val in enumerate(row):
        col = chr(ord('A') + j)
        cell = f"{col}{i+1}"
        ws.Range(cell).Value = val

print(f"Data written: {len(data)} rows x {len(data[0])} cols")

# Format header
ws.Range("A1:D1").Font.Bold = True
ws.Range("A1:D1").Font.Name = "黑体"
ws.Range("A1:D1").Font.Size = 12
ws.Range("A1:D1").Interior.ColorIndex = 15  # gray
print("Header formatted: 黑体 12pt bold")

# Borders
for border_pos in [7,8,9,10]:
    ws.Range("A1:D7").Borders(border_pos).LineStyle = 1
print("Borders applied")

# Formula
ws.Range("C8").Formula = "=SUM(C2:C7)"
ws.Range("A8").Value = "合计"
print(f"Formula set: C8=SUM, Value={ws.Range('C8').Value}")

# Read back
print(f"\nVerification:")
print(f"  A1: {ws.Range('A1').Value}")
print(f"  B3: {ws.Range('B3').Value}")
print(f"  C2: {ws.Range('C2').Value}")
print(f"  C8 (SUM): {ws.Range('C8').Value}")
print(f"  Font A1: {ws.Range('A1').Font.Name} {ws.Range('A1').Font.Size}pt")

# Auto-fit
ws.Range("A:D").Columns.AutoFit()
print("Auto-fit applied")

# Cleanup
wb.Close(False)
app.Quit()

print("\nEXCEL COM VERIFICATION COMPLETE")
