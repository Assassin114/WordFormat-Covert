# DocFormatter — Word 文档格式批量处理工具

基于 Python + PyQt6 的两段式架构：**混乱 Word → Markdown → 格式化 Word**。

## 架构

```
输入 → Word2MD ① → MD 中间态 → MD2Word ② → 输出
         (可审核、可版本管理)
```

## 安装

```bash
cd docformatter
pip install -r requirements.txt
```

## 启动

```bash
python -m docformatter.main
```

## 界面 (2 个 Tab)

| Tab | 功能 |
|-----|------|
| 模板编辑 | 左侧选择项目，右侧配置字体/段落/页面设置 |
| 文档处理 | 3 种模式处理文档 |

### 文档处理 — 3 种模式

| 模式 | 输入 | 模板 | 输出 | 流程 |
|------|------|------|------|------|
| A: Word 格式整理 | .docx | 需要 | .docx | Word→MD→Word 两段式 |
| B: Word→MD | .docx | 不需要 | .md | 仅提取 Markdown |
| C: MD→Word | .md | 需要 | .docx | 从 MD 生成格式化 Word |

## 模板配置项

| 分类 | 项目 | 说明 |
|------|------|------|
| 页面设置 | 纸张大小/方向、页边距、装订线、起始页码 | 对应 Word 页面布局 |
| 封面 | 模板文件 + 占位符映射 + 字体/段落 | {{contract_no}} 等 |
| 签署页 | 模板文件 + 标题 + 字体/段落 | 拟制/校对/标审/审核/批准 |
| 修改记录页 | 表格表头/正文字体段落 | 版本修订记录 |
| 目录页 | 标题/条目字体段落 + 缩进 | 自动生成 TOC 域 |
| 标题1-5 | 字体 + 段落（行距/段前段后/对齐） | 五级标题格式 |
| 正文 | 字体 + 段落（含首行缩进） | 中文国标默认值 |
| 表格 | 基础表/记录表（表头/正文字体段落） | 含背景色 |
| 题注 | 图序/表序/公式序前缀 + 编号格式 | 图/表/公式自动编号 |
| 代码块 | 等宽字体 + 背景色 | 代码段落格式 |
| 页眉页脚 | 封面/目录/正文三类 + 页码格式 | 大写罗马/阿拉伯数字 |

## Markdown 格式规范

### L1 基础（标准 Markdown，用户手写 MD 的最低要求）

```markdown
# 一级标题
## 二级标题

正文段落，**粗体**，*斜体*。

![架构图](images/arch.png)

| 参数 | 说明 |
|------|------|
| 端口 | 8080 |

- 列表项

```python
def hello():
    print("world")
```

内联公式 $E=mc^2$

$$
F = ma
$$
```

系统自动完成：图/表/公式编号、标题书签生成、模板字体段落应用。

### L2 增强（可选 YAML Frontmatter）

```yaml
---
cover_enabled: true
cover_fields:
  contract_no: "XXXX"
  organization: "XX单位"
  date: "2026-04"
signature_enabled: true
revision_enabled: true
toc_enabled: true
---
```

### L3 精确（全功能，用于 Word→MD→Word 管道）

```markdown
<!-- section: cover -->
<!-- section: body -->

## 二级标题  {#sec:xxx}

![架构图](images/arch.png)
*图 1 架构图* {#fig:arch}

*表 1 参数列表* {#tbl:params}
| 参数 | 值 |
|------|-----|
| A    | 1  |

$$
E = mc^2
$$ {#eq:einstein}

如[图 1](#fig:arch)所示...         → Word REF 域（可点击跳转）
参考：[Google](https://google.com)  → Word HYPERLINK

正文内容[^1]                       → Word 脚注
[^1]: 这是脚注的详细内容。

<!-- visio: attachments/x.vsd -->   → Visio OLE 嵌入
```

### 列表编号

```markdown
A. 项目一      → A/B/C 字母编号
B. 项目二
   a. 子项一   → a/b/c 字母编号
```

### MD 语法 → Word 元素对照

| MD 语法 | Word 元素 |
|---------|----------|
| `# ~ #####` | Heading 段落 |
| `**text**` / `*text*` | 粗体 / 斜体 |
| `` `code` `` | 等宽字体 run |
| ` ```语言 代码块 ``` ` | 代码块 + 底纹 |
| `![alt](path)` + 题注 | 图片 + 题注段落 + bookmark |
| `*表注*` + 表格 | 题注段落 + Word table |
| `$$...$$` | LaTeX → PNG → 嵌入 |
| `[text](#fig:xxx)` | REF 域 → bookmark（可点击跳转） |
| `[text](#tbl:xxx)` | REF 域 → 表格跳转 |
| `[text](#eq:xxx)` | REF 域 → 公式跳转 |
| `[text](#sec:xxx)` | REF 域 → 标题跳转 |
| `[text](https://...)` | HYPERLINK 超链接 |
| `[^id]` | FOOTNOTE 脚注 |
| `{#id}` | Word bookmark 锚点 |
| `<!-- section: xxx -->` | Word 分节符 |
| `<!-- visio: path -->` | Visio OLE 嵌入 |

### 降级策略

| 缺失项 | 行为 |
|--------|------|
| 无 YAML frontmatter | 使用模板默认值，全部段落为 body |
| 无 `<!-- section -->` | 整个文档为一个节 |
| 无 `{#id}` 书签 | 自动生成 `fig:N` / `tbl:N` / `eq:N` |
| `[text](#id)` 指向不存在的 bookmark | 降级为纯文本 |
| 图片路径不存在 | 插入占位符文本 + 警告 |
| LaTeX 渲染失败 | 插入原始 LaTeX + 降级提示 |

## 文件结构

```
docformatter/
├── main.py                          # 程序入口
├── models/
│   └── template_config.py           # 数据模型（TemplateConfig 等）
├── core/
│   ├── word2md_converter.py         # Word→MD 转换器
│   ├── md2word_converter.py         # MD→Word 转换器
│   ├── document_analyzer.py         # Word 文档结构分析
│   ├── header_footer.py             # 页眉页脚管理器
│   ├── cover_replacer.py            # 封面占位符替换
│   ├── table_handler.py             # 表格格式化
│   ├── numbering.py                 # 图/表/公式编号管理 (@deprecated)
│   ├── toc_generator.py             # 目录生成 (@deprecated)
│   ├── cross_reference.py           # 交叉引用管理 (@deprecated)
│   ├── formatter.py                 # 旧格式化引擎 (@deprecated)
│   └── style_mapper.py              # 样式映射 (@deprecated)
├── gui/
│   ├── main_window.py               # 主窗口（2 个 Tab）
│   ├── template_config.py           # 模板编辑 Tab
│   └── doc_process.py               # 文档处理 Tab（3 种模式）
├── templates/
│   └── template_io.py               # 模板序列化/反序列化
├── utils/
│   ├── font_formatter.py            # 统一字体/段落格式化
│   ├── placeholder_scanner.py        # 占位符扫描
│   └── logger.py                    # 日志
└── requirements.txt
```

## 依赖

```
python-docx >= 0.8
PyQt6 >= 6.6
markdown-it-py >= 3.0
matplotlib >= 3.8
PyYAML >= 6.0
lxml
```

## 运行测试

```bash
# Word→MD
python -c "from docformatter.core.word2md_converter import convert_word_to_markdown; convert_word_to_markdown('sample.docx')"

# MD→Word
python -c "from docformatter.core.md2word_converter import MD2WordConverter; MD2WordConverter().convert('sample.md', 'output.docx')"
```
