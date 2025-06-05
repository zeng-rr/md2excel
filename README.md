# md2excel

Markdown Table to Excel Converter

[![PyPI version](https://img.shields.io/pypi/v/md2excel)](https://pypi.org/project/md2excel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features
- Automatically recognize table structures in Markdown
- Support multi-table processing with separate sheets in Excel
- Built-in style configurations with customization options (column width, row height, borders, header fonts, cell alignment)
- Command-line interface support

## Installation
```bash
pip install git+https://github.com/zeng-rr/md2excel.git
```

## Usage Example
```python
from md2excel import MarkdownToExcelConverter

converter = MarkdownToExcelConverter()
converter.convert(md_content, "output.xlsx")
```

## Command Line Usage
```bash
md2excel input.md output.xlsx
```

## Contributing
Welcome to submit issues and PRs!

## Documentation
Chinese version available at [README_zh.md](README_zh.md)