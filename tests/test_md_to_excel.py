from pathlib import Path
import sys
import unittest
sys.path.insert(0, str(Path(__file__).parent.parent))
from src.md2excel import MarkdownToExcelConverter

class TestMarkdownToExcel(unittest.TestCase):
    def test_conversion(self):
        md_content = """
# 销售报告

## 月度销售

| 产品 | 销量 | 销售额 |
|------|-----|-------|
| A    | 100 | 5000  |
| B    | 200 | 8000  |

## 客户统计

| 客户名称 | 地区 | 订单数 |
|----------|------|--------|
| 客户A    | 北京 | 5      |
| 客户B    | 上海 | 8      |
    """
        converter = MarkdownToExcelConverter()
        converter.convert(md_content, 'test_output.xlsx')
        
if __name__ == '__main__':
    unittest.main()