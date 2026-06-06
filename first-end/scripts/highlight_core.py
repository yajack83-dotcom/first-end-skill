#!/usr/bin/env python3
"""
核心代码加粗后处理脚本

处理 pandoc 生成的 HTML：
  1. 解析 HTML，找到每个 <h3> 知识点标题
  2. 根据知识点名匹配核心行正则模式
  3. 在每个核心行内，找到 pandoc 生成的 <span class="XX"> 标签
  4. cf/kw/pp/bu/im 类 span → 包 <b><u>（固定骨架）
  5. st/va/dv/co 类 span → 不包（会变的东西）
  6. op 类 span（运算符）→ 不包（题目决定）
  7. 裸文本中的 : ( ) → 包 <b><u>（格式符号）

用法: python highlight_core.py output.html
"""

import re
import sys
import os
from html.parser import HTMLParser


# ── 固定骨架的 pandoc span class ──────────────────────────────
FIXED_CLASSES = {"cf", "kw", "pp", "bu", "im"}  # ControlFlow, Keyword, Preprocessor, BuiltIn, Import

# ── 格式符号（裸文本中需要加粗）────────────────────────────────
FORMAT_CHARS = {":", "(", ")", "[", "]", "{", "}", ";", ",", "."}

# ── 知识点 → 核心行正则 ────────────────────────────────────────
# 每个知识点对应一个正则列表，匹配该知识点的核心代码行
CORE_PATTERNS = {
    "引发异常": [r"\braise\b"],
    "raise": [r"\braise\b"],
    "try…except…else…finally": [r"\btry\b", r"\bexcept\b", r"\belse\b", r"\bfinally\b"],
    "try": [r"\btry\b", r"\bexcept\b", r"\belse\b", r"\bfinally\b"],
    "捕获异常": [r"\bexcept\b"],
    "捕获异常的顺序": [r"\bexcept\b"],
    "finally块": [r"\bfinally\b", r"\bbreak\b"],
    "finally": [r"\bfinally\b"],
    "自定义异常类": [r"\bclass\b", r"\bdef\s+__init__\b", r"\bdef\s+__str__\b", r"\braise\b"],
    "assert语句": [r"\bassert\b"],
    "assert": [r"\bassert\b"],
    "print跟踪法": [r"\bprint\b"],
    "默认配置": [r"\blogging\.(debug|info|warning|error|critical)\b"],
    "basicConfig": [r"\blogging\.basicConfig\b"],
    "输入整数验证": [r"\bdef\b", r"\btry\b", r"\bexcept\b", r"\belse\b", r"\bbreak\b", r"\breturn\b"],
    "文件异常处理": [r"\btry\b", r"\bwith\b", r"\bexcept\b", r"\belse\b", r"\bbreak\b"],
    # 通用：任何含异常处理关键字的代码块
    "异常处理": [r"\btry\b", r"\bexcept\b", r"\belse\b", r"\bfinally\b", r"\braise\b"],
    "自定义异常": [r"\bclass\b.*\bException\b", r"\braise\b"],
    "函数定义": [r"\bdef\b", r"\breturn\b"],
    "类定义": [r"\bclass\b", r"\bdef\s+__init__\b"],
    "循环": [r"\bfor\b", r"\bwhile\b", r"\bbreak\b", r"\bcontinue\b"],
    "条件判断": [r"\bif\b", r"\belif\b", r"\belse\b"],
    "文件操作": [r"\bopen\b", r"\bwith\b", r"\bclose\b"],
    "列表解析": [r"\bfor\b.*\bin\b"],
    "字典解析": [r"\bfor\b.*\bin\b"],
    "集合解析": [r"\bfor\b.*\bin\b"],
}


def match_core_patterns(title_text):
    """根据知识点标题返回匹配的正则列表"""
    patterns = []
    title_lower = title_text.lower().strip()
    for key, pats in CORE_PATTERNS.items():
        if key.lower() in title_lower:
            patterns.extend(pats)
    return list(set(patterns))  # 去重


def wrap_fixed_spans(html_content):
    """在 HTML 中，把固定骨架的 span 标签包上 <b><u>，格式符号裸文本也包上"""

    # Step 1: 找到所有 <h3> 知识点标题及其后续内容范围
    # 简化策略：找到 h3 标签，确定标题文本，然后在后续代码块中应用规则

    # 收集所有 h3 位置
    h3_positions = []
    for m in re.finditer(r'<h3[^>]*>(.*?)</h3>', html_content, re.DOTALL):
        title = re.sub(r'<[^>]+>', '', m.group(1)).strip()
        h3_positions.append((m.start(), m.end(), title))

    if not h3_positions:
        # 没有 h3 标题，全局处理
        return _process_code_blocks(html_content, [])

    # 确定每个 h3 的范围
    result_parts = []
    last_end = 0

    for i, (start, end, title) in enumerate(h3_positions):
        # h3 之前的部分不变
        result_parts.append(html_content[last_end:start])

        # h3 标签本身
        result_parts.append(html_content[start:end])

        # 这个 h3 的范围：到下一个 h3 或文档结束
        next_start = h3_positions[i + 1][0] if i + 1 < len(h3_positions) else len(html_content)
        section_html = html_content[end:next_start]

        # 获取这个知识点的核心正则
        patterns = match_core_patterns(title)

        # 处理这段区域的代码块
        section_html = _process_code_blocks(section_html, patterns)
        result_parts.append(section_html)

        last_end = next_start

    # 最后一个 h3 之后的内容
    if h3_positions:
        result_parts.append(html_content[h3_positions[-1][1]:])

    return "".join(result_parts)


def _process_code_blocks(html, patterns):
    """处理 HTML 中的代码块，对核心行加粗"""

    # 找到所有 <pre><code> 或 <code> 块
    def process_code_block(match):
        code_block = match.group(0)
        pre_tag = ""
        code_tag = ""
        content = code_block

        # 提取 pre/code 标签
        pre_match = re.match(r'(<pre[^>]*>)', code_block)
        code_open_match = re.search(r'(<code[^>]*>)', code_block)
        code_close = '</code>'
        pre_close = '</pre>'

        if pre_match:
            pre_tag = pre_match.group(1)
        if code_open_match:
            code_tag = code_open_match.group(1)

        # 提取 code 内部内容
        inner_match = re.search(r'<code[^>]*>(.*?)</code>', code_block, re.DOTALL)
        if not inner_match:
            return code_block
        inner = inner_match.group(1)

        # 逐行处理
        lines = inner.split('\n')
        processed_lines = []

        for line in lines:
            processed_lines.append(_process_core_line(line, patterns))

        new_inner = '\n'.join(processed_lines)

        # 重建代码块
        if pre_tag and code_tag:
            return f'{pre_tag}{code_tag}{new_inner}{code_close}{pre_close}'
        elif code_tag:
            return f'{code_tag}{new_inner}{code_close}'
        return new_inner

    # 匹配 <pre>...</pre> 包含 <code> 的块
    html = re.sub(
        r'<pre[^>]*>.*?<code[^>]*>.*?</code>.*?</pre>',
        process_code_block,
        html,
        flags=re.DOTALL,
    )

    return html


def _process_core_line(line, patterns):
    """判断一行是否是核心行，如果是则对固定部分加粗"""

    # 去掉 HTML 标签检查文本内容
    text_only = re.sub(r'<[^>]+>', '', line).strip()

    # 空行或纯注释行不处理
    if not text_only:
        return line

    # 检查是否匹配核心正则
    is_core = False
    if patterns:
        for pat in patterns:
            if re.search(pat, text_only):
                is_core = True
                break
    else:
        # 没有指定 patterns，默认加粗常见关键字
        if re.search(r'\b(try|except|finally|else|if|elif|for|while|def|class|return|raise|break|continue|with|assert|import|from|as|pass|yield|lambda|global|nonlocal)\b', text_only):
            is_core = True

    if not is_core:
        return line

    # 处理这一行：包裹固定 span 和格式符号
    return _wrap_fixed_parts(line)


def _wrap_fixed_parts(line):
    """在线内：固定 span class → 包 <b><u>，格式符号裸文本 → 包 <b><u>"""

    # 1. 处理 pandoc span：固定 class 的包 <b><u>
    def wrap_span(match):
        full = match.group(0)
        classes_str = match.group(1)
        inner = match.group(2)

        classes = set(classes_str.split())
        if classes & FIXED_CLASSES:
            return f'<b><u><span class="{classes_str}">{inner}</span></u></b>'
        return full

    line = re.sub(
        r'<span class="([^"]*)">(.*?)</span>',
        wrap_span,
        line,
        flags=re.DOTALL,
    )

    # 2. 处理裸文本中的格式符号（不在任何 HTML 标签内的）
    # 简化策略：在 >...< 之间的文本中替换
    def wrap_bare_chars(match):
        text = match.group(0)
        # 只处理不在标签内的格式符号
        for ch in FORMAT_CHARS:
            text = text.replace(ch, f'<b><u>{ch}</u></b>')
        return text

    # 匹配标签之间的文本
    line = re.sub(
        r'(?<=>)[^<]+(?=<)',
        wrap_bare_chars,
        line,
    )

    return line


def main():
    if len(sys.argv) < 2:
        print("用法: python highlight_core.py <html_file>", file=sys.stderr)
        sys.exit(1)

    html_path = sys.argv[1]
    if not os.path.exists(html_path):
        print(f"文件不存在: {html_path}", file=sys.stderr)
        sys.exit(1)

    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    processed = wrap_fixed_spans(html_content)

    # 覆盖原文件
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(processed)

    print(f"[highlight_core] ✅ 处理完成: {html_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
