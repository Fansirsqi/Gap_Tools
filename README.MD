## Gap with Excel file

### 1. How To Start

```bash
git clone https://github.com/Fansirsqi/Gap_Tools.git

cd Gap_tools


python -m venv .venv

# 指定清华源安装依赖
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

```

格式化文件

pip install yapf

`yapf --style="{column_limit: 280}" -i .\gaps.py`

`yapf --style="{column_limit: 180, indent_closing_brackets: True}" -i .\gaps.py`
