# first-end

**期末复习梳理技能** — Claude Code 自定义技能，给一个课件文件夹，自动生成结构化期末复习文档。

## 做什么

1. **概念/实记内容** → 术语 + 大白话 + 记什么（覆盖选择题、判断题）
2. **代码阅读** → 三步过滤（PPT 代码 → 归属知识点 → 分析结构类型）→ 每行加一针见血的注释 + 出题方式标注（覆盖大题）

**不出题，只梳理。**

## 适用场景

4 种考试题型的应对：

| 题型 | 形式 | 文档如何覆盖 |
|:---|:---|:---|
| ① | 写出解析式的结果 | 解析式代码 + 逐行注释 + 运行结果 |
| ② | 执行程序写出运行结果 | 完整程序逐行追踪注释 |
| ③ | 写出代码运行结果 | 代码片段逐行注释 |
| ④ | 补全代码 | 关联结构标注（try→except、def→调用等） |

## 安装

1. 将整个 `first-end/` 文件夹放入 `~/.claude/skills/` 目录
2. 安装 Python 依赖：

```bash
pip install pdfplumber python-pptx python-docx
```

> 旧格式 `.ppt` / `.doc` / `.xls` 需要安装 Microsoft Office，脚本通过 Windows COM 自动转换后提取。

## 使用

在 Claude Code 中输入：

```
/first-end
```

然后提供课件文件夹路径（包含 PPT/PDF/教材），技能会自动：

1. 提取所有文件文本（支持 100+ 格式）
2. 区分概念型 vs 代码型内容
3. 概念 → 3 列表格（术语 | 大白话 | 记什么）
4. 代码 → 三步过滤 → 公共依赖 + 逐行注释 + 出题标注
5. 生成自包含 HTML（双击浏览器打开，可 Ctrl+P 打印 PDF）

## 文件结构

```
first-end/
├── skill.md                  # 技能定义
├── scripts/
│   ├── extract_text.py       # 全格式文本提取（PPT/PDF/DOCX 等）
│   ├── highlight_core.py     # HTML 核心代码加粗后处理
│   └── style-fix.html        # pandoc HTML 样式（表格/代码块/打印）
└── README.md
```

## 核心方法论：三步过滤

```
PPT 全部代码
  ↓ 第一步：归属知识点（这段代码属于哪个章节的哪个知识点？）
  ↓ 第二步：分析结构类型（单独块？固定格式闭环？函数分散？）
  ↓ 第三步：按类型输出（每行注释 + 出题方式标注）
```

只保留能对应 4 种考试题型的代码，纯演示/装饰/操作步骤/思政案例一律丢弃。

## 依赖

- Python 3.8+
- pandoc（HTML 转换）
- `pip install pdfplumber python-pptx python-docx`
- Windows COM（旧格式 `.ppt`/`.doc` 需要 Office）

## License

MIT
