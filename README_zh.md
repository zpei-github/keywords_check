# PDF 关键字搜索工具

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

PDF 关键字搜索工具，支持跨行/跨页句子提取、噪声自动过滤、多格式导出。

## 适用场景

本工具专为需要从 PDF 文档中**快速定位关键信息**的场景设计：

- **招标文件** — 搜索"公章""授权""资质""业绩"等硬性要求，自动计算重要性得分，避免遗漏关键条款
- **法律合同** — 快速检索"违约责任""知识产权""保密""赔偿"等核心条款，跨页句子自动拼接
- **教材课本** — 按章节搜索知识点关键字，提取上下文形成复习提纲
- **技术文档** — 检索 API 名称、配置参数、错误码等关键信息
- **政府公文** — 搜索"批复""决定""意见"等关键词，自动过滤红头文件的页眉页脚

## 核心功能

- **多关键字加权搜索** — 每个关键字可设置独立的重要性分数（例如"承诺"9 分、"签字"3 分），结果按得分排序
- **完整句子提取** — 自动返回关键字所在的完整句子，上下文窗口可配置
- **跨行/跨页处理** — 正确处理被 PDF 换行符切断的关键字，以及跨越页边界的句子
- **噪声自动检测** — 智能识别并过滤页眉、页脚、页码、水印等重复内容
- **关键字保护** — 包含目标关键字的文本块即使在边缘区域也不会被过滤
- **多格式导出：**
  - **Excel** — 按重要性排名，关键字红色高亮，支持跨页标记
  - **PDF** — 在原 PDF 上高亮关键字（含跨行/跨页高亮）
  - **TXT** — 完整分析报告，含命中统计和噪声检测详情
- **PySide6 图形界面** — 可视化操作，无需编写代码

## 安装

```bash
git clone https://github.com/zpei-github/pdf-keyword-finder.git
cd pdf-keyword-finder
pip install -r requirements.txt
```

### 依赖

| 包名 | 版本要求 |
| ---- | -------- |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | >= 1.23.0 |
| [openpyxl](https://openpyxl.readthedocs.io/) | >= 3.1.0 |
| [PySide6](https://doc.qt.io/qtforpython/) | >= 6.5, < 7.0 |

## 快速上手

### 命令行

```python
from pdf_keyword_finder import find_keywords_in_pdf

# 招标文件关键字示例：关键字 → 重要性分数
keywords = {
    "公章": 7,
    "鲜章": 7,
    "承诺": 9,
    "授权": 6,
    "证明": 9,
    "证明材料": 9,
    "签字": 3,
    "盖章": 7,
    "必须": 4,
}

results = find_keywords_in_pdf(
    pdf_path="招标文件.pdf",
    keywords=keywords,
    context_rich=100,       # 句子上下文扩展字符数
    front_window=0,         # 向前搜索窗口
    output_file="分析报告.txt",       # 可选
    excel_file="搜索结果.xlsx",       # 可选
    highlight_pdf="高亮版.pdf",       # 可选
    auto_clean_noise=True,            # 自动过滤页眉页脚
    header_ratio=0.15,                # 页眉区域占比
    footer_ratio=0.85,                # 页脚区域占比
    repeat_threshold=0.3,             # 噪声重复率阈值
)
```

### 图形界面

```bash
python gui.py
```

GUI 提供文件选择、关键字输入、噪声过滤参数调节、结果浏览等可视化操作。

## API 参考

### `find_keywords_in_pdf()`

主入口函数。

| 参数 | 类型 | 默认值 | 说明 |
| ---- | ---- | ------ | ---- |
| `pdf_path` | `str` | 必填 | PDF 文件路径 |
| `keywords` | `List[str] \| Dict[str, int]` | 必填 | 关键字列表或 关键字→分数 字典 |
| `context_rich` | `int` | 必填 | 向前/向后扩展句子边界的最大字符数 |
| `front_window` | `int` | 必填 | 从关键字位置向前搜索的最大字符数（上限 80） |
| `output_file` | `str \| None` | `None` | TXT 分析报告输出路径 |
| `excel_file` | `str \| None` | `None` | Excel 结果输出路径 |
| `highlight_pdf` | `str \| None` | `None` | 高亮 PDF 输出路径 |
| `auto_clean_noise` | `bool` | `False` | 是否启用页眉/页脚/页码过滤 |
| `header_ratio` | `float` | `0.15` | 页眉检测区域占比（页面顶部） |
| `footer_ratio` | `float` | `0.85` | 页脚检测区域占比（页面底部） |
| `repeat_threshold` | `float` | `0.3` | 判定为噪声的最低重复率 |

**返回值** 字典包含：

- `total_matches` — 匹配到的句子总数
- `by_page` — 按页码分组的搜索结果
- `all_results` — 所有结果的平铺列表
- `noise_info` — 被过滤的噪声块详情

## 工作原理

1. **文本提取** — PyMuPDF 提取 PDF 文本块及其坐标和页码元数据
2. **噪声检测** — 位于页面边缘且重复率高的文本块被标记为噪声；包含关键字的块始终保留
3. **关键字匹配** — 所有关键字合并为单个正则模式，一次遍历完成全文搜索
4. **句子提取** — 从每个匹配位置向外扩展，按标点符号确定句子边界
5. **重叠合并** — 重叠句子自动合并，关键字集合汇总
6. **页码映射** — 通过前缀和数组将字符位置映射回原始页码（O(log n) 复杂度）
7. **结果导出** — 按需生成 TXT / Excel / 高亮 PDF

## 性能优化

- **单次正则匹配** — 所有关键字合并为一个编译后的正则模式
- **列表拼接** — 使用 list + join（O(n)）替代字符串累加（O(n²)）
- **前缀和页码查找** — 二分查找实现 O(log n) 的页码定位
- **逐页缓存** — 高亮生成时每页只调用一次 `get_text`，缓存后复用

## License

MIT
