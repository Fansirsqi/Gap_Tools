from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill,Alignment

rules = []
"""规则列表"""

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

center_align = Alignment(horizontal='center', vertical='center')
"""居中对其"""

if __name__ == "__main__":
    pass
    # for rule in rules:
    # ws.conditional_formatting.add("C1:C38", rule)
