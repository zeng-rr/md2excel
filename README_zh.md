# md2excel

Markdown表格转Excel转换工具

[![PyPI version](https://img.shields.io/pypi/v/md2excel)](https://pypi.org/project/md2excel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 功能特性
- 自动识别Markdown中的表格结构
- 支持对单个markdown文件多表格分Sheet存储并能自动匹配每个表格的标题
- 自带基础样式设置并能够自定义样式配置（列宽、行高、边框、表头字体、单元格字体和对齐等）
- 支持命令行转换

## 安装
```bash
pip install git+https://github.com/zeng-rr/md2excel.git
```

## 使用示例
```python
from md2excel import MarkdownToExcelConverter

converter = MarkdownToExcelConverter()
converter.convert(md_content, "output.xlsx")
```

## 命令行使用
```bash
md2excel input.md output.xlsx
```

## 贡献指南
欢迎提交Issue和PR！