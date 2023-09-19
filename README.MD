# Gap with Excel file

## 1. 如何开始

```bash
git clone https://github.com/Fansirsqi/Gap_Tools.git

cd Gap_tools

python -m venv .venv

# 指定清华源安装依赖
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## 2. 单元格`数据`格式设置指南

```python
from openpyxl import Workbook
import datetime
path = r'xxxx/xxxx/xxxx'
wb = Workbook()
ws = wb.active
ws['A2'] = 1
ws['A2'].number_format = 'General'  # 常规格式

ws['A3'] = 1
ws['A3'].number_format = '0.00'  # 保留两位小数

ws['A3'] = 1
ws['A3'].number_format = '0.0%'  # 设置成百分比格式

ws['A4'] = 100000
ws['A4'].number_format = '#,##0'  # 设置千分位格式

ws['A5'] = 100000
ws['A5'].number_format = '0.00E+00'  # 科学计数法

ws['A6'] = datetime.datetime(2018, 3, 20, 10, 20, 39)
ws['A6'].number_format = 'mm-dd-yy'  # 日期格式

ws['A7'] = datetime.datetime(2018, 3, 20, 10, 20, 39)
ws['A7'].number_format = 'h:mm:ss'  # 时间格式
wb.save(path)

```

## 3. 单元格`规则`设置指南

```py
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

rules = []
# 创建DataBarRule规则
rule1 = CellIsRule(operator=">", formula=[0.1], fill=PatternFill(end_color="00EE822F", fill_type="solid"))
"""大于10%"""
rules.append(rule1)
rule2 = CellIsRule(operator="between", formula=[0.005, 0.1], fill=PatternFill(end_color="00FFEFC1", fill_type="solid"))
"""介于0.5%到10%"""
rules.append(rule2)
rule3 = CellIsRule(operator="between", formula=[-0.1, -0.005], fill=PatternFill(end_color="00FCDECD", fill_type="solid"))
"""介于-0.5%到-10%"""
rules.append(rule3)
rule4 = CellIsRule(operator="<", formula=[-0.005], fill=PatternFill(end_color="00F4B382", fill_type="solid"))
"""小于-10%"""
rules.append(rule4)

for rule in rules:
    ws.conditional_formatting.add("C1:C38", rule)
```
