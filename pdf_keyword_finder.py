"""
PDF关键字搜索工具
功能：
- 支持搜索多个关键字
- 返回关键字所在的完整句子
- 正确处理跨行句子
- 正确处理跨页句子
- 按重要性排序导出Excel

性能优化：
- 使用列表+join代替字符串拼接
- 合并关键字为单个正则模式，一次遍历
- 二分查找页码，O(log n)复杂度
- 预编译正则表达式
"""

import fitz  # PyMuPDF
import re
import random
import bisect
from typing import List, Dict, Tuple
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


# 预编译正则表达式
WHITESPACE_PATTERN = re.compile(r'\s+')
EXTRA_BOUNDARY_CHARS = set(',，;；')
BOUNDARY_CHARS = set('。！？!?')
IGNOR_CHARS = set('(（注')


def auto_detect_noise_blocks(doc, check_num: int) -> Tuple[int, int]:
    """
    自动检测页眉页脚等噪音block数量

    检测逻辑：
    - blocks顺序通常是：正文 -> 页眉页脚 -> 页码/水印（末尾）
    - 开头相同的block判定为页眉（如果页眉出现在blocks开头）
    - 末尾相同的block判定为页脚/水印
    - 最后一个block通常是页码，每页不同，默认跳过1个

    参数:
        doc: fitz.Document对象
        check_num: 检测的页面数量

    返回:
        (skip_start_block, skip_end_block) 需要跳过的开头和末尾block数量
    """
    skip_start_block = 0
    skip_end_block = 0  # 默认跳过最后一个block（通常是页码）

    if len(doc) <= 2 * check_num:
        return skip_start_block, skip_end_block

    # 从中间随机选几页进行比对，避免封面等特殊情况
    random_list = random.sample(range(check_num, len(doc)), check_num)
    page_blocks = [doc[page_num].get_text_blocks() for page_num in random_list]

    # 检测开头的噪音block（页眉）
    # 如果多个页面开头的block文本相同，则为噪音
    while True:
        texts = set()
        valid = True
        for blocks in page_blocks:
            if len(blocks) <= skip_start_block:
                valid = False
                break
            texts.add(blocks[skip_start_block][4])

        if not valid or len(texts) > 1:
            break
        skip_start_block += 1

    # 检测末尾的噪音block（页脚/水印）
    # 先跳过最后一个（页码），然后检查倒数第二个、第三个...
    while True:
        texts = set()
        valid = True
        for blocks in page_blocks:
            idx = -(skip_end_block + 1)  # 倒数第(skip_end_block+1)个
            if len(blocks) <= skip_end_block + 1:
                valid = False
                break
            texts.add(blocks[idx][4])

        if not valid or len(texts) > 1:
            break
        skip_end_block += 1

    return skip_start_block, skip_end_block


def get_page_text_with_layout(
    pdf_path: str,
    check_num: int,
    skip_start_block: int = 0,
    skip_end_block: int = 0,
    auto_clean_noise: bool = True
) -> Tuple[str, List[Dict]]:
    """
    获取PDF文本，将每个block去除换行符后拼接

    优化：使用列表+join代替字符串拼接，时间复杂度从O(n²)降至O(n)

    参数:
        pdf_path: PDF文件路径
        skip_start_block: 忽略每页开头的n个block（手动指定）
        skip_end_block: 忽略每页末尾的n个block（手动指定）
        auto_clean_noise: 是否自动检测并去除页眉页脚页码
        check_num: 自动检测时采样的页面数量

    返回:
        (拼接后的完整文本, block信息列表)
    """
    doc = fitz.open(pdf_path)
    text_parts = []  # 使用列表收集文本，避免O(n²)拼接
    block_info = []  # 记录每个block的信息
    current_pos = 0  # 当前文本位置

    # 自动检测噪音block
    if auto_clean_noise:
        auto_skip_start, auto_skip_end = auto_detect_noise_blocks(doc, check_num)
        skip_start_block = max(skip_start_block, auto_skip_start)
        skip_end_block = max(skip_end_block, auto_skip_end)
        print(f"自动检测: 跳过开头 {skip_start_block} 个block, 跳过末尾 {skip_end_block} 个block")

    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text_blocks()
        # 去除页面的页眉页脚和页码等噪音
        if skip_end_block > 0:
            blocks = blocks[skip_start_block: len(blocks) - skip_end_block]
        else:
            blocks = blocks[skip_start_block:]
        for block in blocks:
            x0, y0, x1, y1, text, block_no, block_type = block

            # 只处理文本块
            if block_type == 0 and text.strip():
                # 去除block内的换行符
                clean_text = text.replace('\n', '').replace('\r', '')
                clean_text = WHITESPACE_PATTERN.sub(' ', clean_text).strip()

                if clean_text:
                    block_info.append({
                        'page': page_num + 1,
                        'block_no': block_no,
                        'text': clean_text,
                        'position': current_pos
                    })
                    text_parts.append(clean_text)
                    current_pos += len(clean_text)

    doc.close()
    full_text = ''.join(text_parts)  # O(n)拼接
    return full_text, block_info

def find_keywords_in_text(
    full_text: str,
    keywords: List[str],
    context_chars: int
) -> List[Dict]:
    """
    在文本中搜索关键字，返回完整句子（去重后）

    优化：合并所有关键字为单个正则模式，一次遍历完成匹配
    """
    if not keywords:
        return []

    # 构建合并的正则模式：(keyword1|keyword2|...)
    # 使用 re.escape 确保特殊字符正确处理
    escaped_keywords = [re.escape(kw) for kw in keywords]
    pattern = re.compile('(' + '|'.join(escaped_keywords) + ')', re.IGNORECASE)

    # 一次遍历收集所有匹配
    all_matches = []
    for match in pattern.finditer(full_text):
        matched_text = match.group()
        # 找到匹配的是哪个关键字（保持原始大小写）
        for keyword in keywords:
            if keyword.lower() == matched_text.lower():
                all_matches.append({
                    'keyword': keyword,
                    'start': match.start(),
                    'end': match.end()
                })
                break

    if not all_matches:
        return []

    # 按位置排序
    all_matches.sort(key=lambda x: x['start'])

    # 提取句子并根据范围重叠进行合并
    sentence_list = []

    for match in all_matches:
        print(match)
        sentence, sentence_start, sentence_end = extract_sentence_from_text(
            full_text, match['start'], match['end'], context_chars
        )
        sentence_list.append({
            'sentence': sentence,
            'keywords': {match['keyword']},
            'sentence_start': sentence_start,
            'sentence_end': sentence_end,
            'position': match['start']
        })

    # 按句子开始位置排序
    sentence_list.sort(key=lambda x: x['sentence_start'])

    # 合并重叠的句子
    merged = []
    for item in sentence_list:
        if not merged:
            merged.append(item)
            continue

        last = merged[-1]
        if item['sentence_end'] <= last['sentence_end']:
            last['sentence'] = WHITESPACE_PATTERN.sub(' ', last['sentence'])
            last['keywords'].update(item['keywords'])
        elif item['sentence_start'] == last['sentence_start']:
            item['keywords'].update(last['keywords'])
            item['sentence'] = WHITESPACE_PATTERN.sub(' ', item['sentence'])
            merged[-1] = item
        else:
            merged.append(item)


    return merged


def extract_sentence_from_text(text: str, start_pos: int, end_pos: int, context_chars: int) -> Tuple[str, int, int]:
    """从文本中提取包含关键字的完整句子"""
    text_len = len(text)

    # 向前找句子开始
    sentence_start = start_pos
    min_bound = max(0, start_pos - context_chars)  # 向前边界

    for i in range(start_pos - 1, -1, -1):
        current_char = text[i]
        # 边界检查：下一字符是否存在
        next_char = text[i + 1] if i + 1 < text_len else ''

        # 在 context_chars 范围内：只检测 BOUNDARY_CHARS
        # 突破边界后：同时检测 BOUNDARY_CHARS 和 EXTRA_BOUNDARY_CHARS
        if i >= min_bound:
            # 范围内：遇到句末标点且下一字符不在 IGNOR_CHARS 中才截断
            if current_char in BOUNDARY_CHARS and next_char not in IGNOR_CHARS:
                sentence_start = i + 1
                break
        else:
            # 突破边界：遇到任何分界符都截断
            if current_char in BOUNDARY_CHARS or current_char in EXTRA_BOUNDARY_CHARS:
                sentence_start = i + 1
                break

        sentence_start = i

    # 向后找句子结束
    sentence_end = end_pos
    max_bound = min(text_len, end_pos + context_chars)  # 向后边界

    for i in range(end_pos, text_len):
        current_char = text[i]

        # 在 context_chars 范围内：只检测 BOUNDARY_CHARS
        # 突破边界后：同时检测 BOUNDARY_CHARS 和 EXTRA_BOUNDARY_CHARS
        if i < max_bound:
            # 范围内：遇到句末标点就截断
            if current_char in BOUNDARY_CHARS:
                sentence_end = i + 1
                break
        else:
            # 突破边界：遇到任何分界符都截断
            if current_char in BOUNDARY_CHARS or current_char in EXTRA_BOUNDARY_CHARS:
                sentence_end = i + 1
                break
        sentence_end = i + 1

    sentence = text[sentence_start:sentence_end].strip()
    # 清理多余空格
    sentence = WHITESPACE_PATTERN.sub(' ', sentence)
    return sentence, sentence_start, sentence_end


def find_keywords_in_pdf(
    pdf_path: str,
    context_width: int,
    auto_check_num: int,
    keywords: List[str] | Dict[str, int],
    output_file: str | None = None,
    excel_file: str | None = None
) -> Dict:
    """
    在PDF中搜索关键字并返回结果

    参数:
        pdf_path: PDF文件路径
        keywords: 关键字列表或字典（关键字->分数）
        output_file: 可选的输出文件路径（txt）
        excel_file: 可选的Excel输出文件路径（按重要性排序）
    """
    print(f"正在打开PDF文件: {pdf_path}")

    # 支持字典格式（关键字+分数）
    if isinstance(keywords, dict):
        keywords_point = keywords
        keywords_list = list(keywords.keys())
    else:
        keywords_list = keywords
        # 默认每字1分
        keywords_point = {k: 1 for k in keywords_list}

    # 获取拼接后的文本
    full_text, block_info = get_page_text_with_layout(pdf_path, check_num=auto_check_num, auto_clean_noise=True)
    print(f"文本总长度: {len(full_text)} 字符, 共 {len(block_info)} 个文本块")

    # 搜索关键字
    print(f"正在搜索关键字: {keywords_list}")
    results = find_keywords_in_text(full_text, keywords_list, context_width)

    print(f"\n找到 {len(results)} 处匹配")

    # 计算每个句子的总分
    for result in results:
        score = sum(keywords_point.get(kw, 1) for kw in result['keywords'])
        result['score'] = score

    # 构建页码位置索引，用于二分查找
    # block_info 已按 position 排序
    page_positions = [block['position'] for block in block_info]
    page_numbers = [block['page'] for block in block_info]

    # 按页码组织结果
    page_results = {}
    for result in results:
        pos = result['position']
        # 使用二分查找快速定位页码
        idx = bisect.bisect_right(page_positions, pos)
        if idx > 0:
            page_num = page_numbers[idx - 1]
        else:
            page_num = page_numbers[0] if page_numbers else 1

        result['page'] = page_num

        if page_num not in page_results:
            page_results[page_num] = []
        page_results[page_num].append(result)

    # 保存到txt文件
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("PDF关键字搜索结果\n")
            f.write("=" * 60 + "\n")
            f.write(f"PDF文件: {pdf_path}\n")
            f.write(f"搜索关键字: {', '.join(keywords_list)}\n")
            f.write(f"总匹配句子数: {len(results)}\n")
            f.write("=" * 60 + "\n\n")

            for page_num in sorted(page_results.keys()):
                f.write(f"--- 第 {page_num} 页 ---\n")
                for idx, r in enumerate(page_results[page_num], 1):
                    keywords_str = ', '.join(r['keywords'])
                    f.write(f"\n[{idx}] 关键字: {keywords_str}\n")
                    f.write(f"    完整句子: {r['sentence']}\n")
        print(f"\n结果已保存到: {output_file}")

    # 导出Excel（按重要性排序）
    if excel_file:
        export_to_excel(results, excel_file, pdf_path, keywords_point)
        print(f"Excel已保存到: {excel_file}")

    return {
        'total_matches': len(results),
        'by_page': page_results,
        'all_results': results
    }


def export_to_excel(results: List[Dict], excel_file: str, pdf_path: str, keywords_point: Dict[str, int]):
    sorted_results = sorted(results, key=lambda x: x.get('score', 0), reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "关键字搜索结果"

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=12, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    wrap_alignment = Alignment(wrap_text=True, vertical='top')

    # ✅ 3.1.5 正确格式：InlineFont
    red_font = InlineFont(color='00FF0000')
    default_font = InlineFont(color='00000000')

    ws.merge_cells('A1:F1')
    ws['A1'] = f"PDF关键字搜索结果 - {pdf_path}"
    ws['A1'].font = Font(bold=True, size=14)

    ws.merge_cells('A2:F2')
    keywords_info = ', '.join([f"{k}({v}分)" for k, v in keywords_point.items()])
    ws['A2'] = f"搜索关键字: {keywords_info}"
    ws['A2'].font = Font(italic=True)

    headers = ['排名', '重要性', '得分', '页码', '包含关键字', '完整句子']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 8
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 80

    max_score = max((r.get('score', 0) for r in sorted_results), default=1)

    for idx, result in enumerate(sorted_results, 1):
        row = idx + 4
        score = result.get('score', 0)

        if max_score > 0:
            ratio = score / max_score
            if ratio >= 0.8:
                importance = "★★★★★"
                importance_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
            elif ratio >= 0.6:
                importance = "★★★★"
                importance_fill = PatternFill(start_color="FFA94D", end_color="FFA94D", fill_type="solid")
            elif ratio >= 0.4:
                importance = "★★★"
                importance_fill = PatternFill(start_color="FFD93D", end_color="FFD93D", fill_type="solid")
            elif ratio >= 0.2:
                importance = "★★"
                importance_fill = PatternFill(start_color="ADE25D", end_color="ADE25D", fill_type="solid")
            else:
                importance = "★"
                importance_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        else:
            importance = "★"
            importance_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        ws.cell(row=row, column=1, value=idx).border = thin_border
        ws.cell(row=row, column=2, value=importance).border = thin_border
        ws.cell(row=row, column=2).fill = importance_fill
        ws.cell(row=row, column=3, value=score).border = thin_border
        ws.cell(row=row, column=4, value=result['page']).border = thin_border
        ws.cell(row=row, column=5, value=', '.join(result['keywords'])).border = thin_border
        ws.cell(row=row, column=5).alignment = wrap_alignment

        # ===================== 完美标红逻辑 =====================
        sentence = result['sentence']
        keywords = result.get('keywords', [])
        sentence_cell = ws.cell(row=row, column=6)

        if not keywords:
            sentence_cell.value = sentence
        else:
            parts = []
            current = 0
            pattern = '|'.join(re.escape(k) for k in keywords)
            for match in re.finditer(pattern, sentence):
                s, e = match.span()
                if s > current:
                    parts.append(TextBlock(default_font, sentence[current:s]))
                parts.append(TextBlock(red_font, sentence[s:e]))
                current = e
            if current < len(sentence):
                parts.append(TextBlock(default_font, sentence[current:]))
            
            sentence_cell.value = CellRichText(parts)
        # ========================================================

        sentence_cell.border = thin_border
        sentence_cell.alignment = wrap_alignment

        line_count = max(1, len(sentence) // 60 + 1)
        ws.row_dimensions[row].height = min(60, line_count * 15)

    ws.freeze_panes = 'A5'
    wb.save(excel_file)


# 使用示例
if __name__ == "__main__":
    # 示例：搜索单个PDF文件
    pdf_path = r"E:\Desktop\test.pdf"  # 替换为你的PDF文件路径

    # 定义要搜索的关键字及分数
    keywords_point = {
            "提供": 4,
            "提交": 4,
            "递交": 4,
            "出具": 4,
            "响应": 4,
            "加盖": 5,
            "承诺": 9,
            "授权": 6,
            "证明": 9,
            "公章": 7,
            "鲜章": 7,
            "报告": 5,
            "签字": 3,
            "说明": 3,
            "证书": 4,
            "盖单位章": 7,
            "盖章": 7,
            "必须": 4}

    # 执行搜索（直接传字典，会自动计算分数）
    results = find_keywords_in_pdf(
        pdf_path=pdf_path,
        keywords=keywords_point,
        context_width=200,
        auto_check_num=3,
        output_file=r"E:\Desktop\output.txt",  # 可选：保存txt结果
        excel_file=r"E:\Desktop\output.xlsx"   # 可选：保存Excel结果（按重要性排序）
    )
