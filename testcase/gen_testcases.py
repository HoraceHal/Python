"""根据 cases JSON 和 template.xlsx 生成测试用例 Excel。
用法: python3 testcase/gen_testcases.py <cases.json> <输出文件名.xlsx> [sheet名]
cases.json: [{"name","priority","pre","data","steps","expect"}, ...]
固定列(取自模板): 所属模块/用例类型=功能/测试方式=手动/用例状态=有效; 来源ID、来源名称、标签留空。
"""
import sys, json, copy, os, openpyxl
here = os.path.dirname(os.path.abspath(__file__))
MODULE = '2026年联通在线5G宽视界产品及平台能力研发项目'
cases = json.load(open(sys.argv[1], encoding='utf-8'))
out = sys.argv[2]
wb = openpyxl.load_workbook(os.path.join(here, 'template.xlsx'))
ws = wb.active
if len(sys.argv) > 3: ws.title = sys.argv[3]
sty = [copy.copy(ws.cell(2, c)._style) for c in range(1, 14)]
h = ws.row_dimensions[2].height
ws.delete_rows(2, ws.max_row)
for i, c in enumerate(cases, 2):
    row = [c['name'], MODULE, '功能', c.get('priority', '中'), '手动', '有效', '', '', '',
           c['pre'], c.get('data', ''), c['steps'], c['expect']]
    for j, v in enumerate(row, 1):
        ws.cell(i, j, v)._style = copy.copy(sty[j - 1])
    if h: ws.row_dimensions[i].height = h
os.makedirs(os.path.join(here, 'output'), exist_ok=True)
path = os.path.join(here, 'output', out)
wb.save(path); print(len(cases), '->', path)
