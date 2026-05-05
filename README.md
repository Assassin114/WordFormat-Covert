# DocFormatter

基于 Python + PyQt6 的 Word 文档格式处理工具，采用 **Word → Markdown → Word** 两段式转换架构。

## 功能概述

| 功能 | 说明 |
|------|------|
| 格式整理 | 将格式不统一的 .docx 批量整理为规范化文档 |
| 模板管理 | 可视化编辑格式模板（字体、段落、页面、标题层级、页眉页脚等），支持导入导出共享 |
| Word → Markdown | 将 Word 文档结构化提取为 Markdown（含书签、REF 域、脚注、Visio OLE） |
| Markdown → Word | 将 Markdown 文档生成为格式化 Word（含封面、签署页、目录、图表编号、交叉引用） |
| 多级标题编号 | 支持中文/数字/罗马/字母编号风格，可配置多级继承断链 |
| 页面设置 | 纸张预设（A4/A3/A5/B5/Letter/Legal/自定义）、页边距、装订线、页码范围 |
| 实时预览 | 控件变更后即时渲染预览效果，预览比例跟随页面设置自适应 |

## 安装

### 方式一：使用预编译版本（推荐）

无需安装 Python 或任何依赖。

1. 从发布页面或项目成员处获取 `DocFormatter.zip`
2. 解压到任意目录
3. 双击 `DocFormatter.exe` 启动

### 方式二：从源码运行

```bash
# 克隆仓库
git clone <repo-url>
cd WordFormat-Covert

# 安装依赖
cd docformatter
pip install -r requirements.txt

# 启动程序
cd ..
python -m docformatter.main
```

**自行打包：**

```bash
# Windows
build.bat

# Linux / macOS
bash build.sh
```

打包产物位于 `dist/DocFormatter/`，将该目录打包为 zip 即可分发。

## 界面结构

主窗口包含两个标签页：

| 标签页 | 用途 |
|--------|------|
| 模板编辑 | 可视化配置格式模板（字体、段落、页面设置、标题层级等） |
| 文档处理 | 选择处理模式，加载文件，执行转换 |

工作流程：**先在模板编辑中完成格式配置，再在文档处理中选择模式并执行。**

---

## 使用指南

### 一、模板编辑

界面采用左-中-右三栏布局：

- **左侧** — 配置项目列表，按文档结构从上到下排列（页面设置 → 页眉页脚 → 封面 → 签署页 → 修改记录页 → 目录页 → 标题 → 正文 → 表格 → 题注 → 代码块）
- **中间** — 选中项目的编辑表单
- **右侧** — 风格预览面板，实时显示当前配置的渲染效果

#### 工具栏

| 控件 | 功能 |
|------|------|
| 新建模板 | 重置为系统默认模板 |
| 模板下拉框 | 列出所有可用模板（内置模板 + 用户模板），选择即加载 |
| 保存 | 将当前配置保存到当前模板文件 |
| 另存为 | 将模板保存到用户模板目录，需填写名称、标签、来源文档 |
| 导出 | 将当前模板导出为 `.json` 文件，用于分享给他人 |
| 导入 | 从外部 `.json` 文件导入模板到用户模板目录 |
| 元信息栏 | 编辑模板名称、标签（逗号分隔）、来源文档列表 |

#### 配置项说明

**页面设置**

| 属性 | 说明 |
|------|------|
| 预设 | A4 / A3 / A5 / B5 / Letter / Legal / 自定义 |
| 方向 | 纵向 / 横向 |
| 页边距 | 上、下、左、右（单位：mm） |
| 装订线 | 装订线宽度及位置（左/上） |
| 起始页码 | 文档起始页码 |

**页眉页脚**
- 打印模式：单面打印 / 双面打印
- 公式格式：保持原样 / 转为 OMML / 渲染为图片
- 按页面类型（封面页、目录页、正文页）分别设置：页眉文本、页脚显隐、页码样式（阿拉伯数字 / 大写罗马 / 小写罗马 / 大写字母 / 小写字母）、页码位置及对齐方式
- 首页不同 / 奇偶页不同

**封面 / 签署页 / 修改记录页 / 目录页**

| 页面 | 配置项 |
|------|--------|
| 封面 | 启用开关、模板文件（支持 `{{占位符}}` 替换）、标题字体段落 |
| 签署页 | 启用开关、模板文件、标题文本（自动检测封面标题）、标题字体段落 |
| 修改记录页 | 启用开关、表头字体段落、表体字体段落 |
| 目录页 | 启用开关、标题字体段落、条目字体段落、缩进值 |

**标题（多级列表）**

以表格形式配置五级标题：

| 级别 | 编号格式 | 包含上级 | 编号预览 |
|------|---------|---------|---------|
| 1 | 一、二、三 | 是 | 一、 |
| 2 | 1. / 2. | 是 | 一、1. |
| 3 | (1) / (2) | 是 | 一、1.(1) |
| 4 | 无编号 | 否 | (无) |

- 选中某行后，下方编辑区可配置该级别的字体和段落属性
- "包含上级"勾选后，下级编号将继承上级前缀（类似 Word 的多级列表逻辑）
- 某级别取消"包含上级"后，其上各级的编号链将在此断开
- "编号预览"列根据各行的编号格式和包含上级状态实时计算

**正文 / 表格 / 题注 / 代码块**

| 模块 | 配置项 |
|------|--------|
| 正文 | 中英文字体、字号、首行缩进、行距、段前段后、对齐方式 |
| 表格 | 分基础表/记录表，分别配置表头/表体的字体段落及表头背景色 |
| 题注 | 图/表/公式前缀，题注位置，编号模式（全局统一 / 按类型独立） |
| 代码块 | 等宽字体及背景色 |

---

### 二、文档处理

#### 处理模式

| 模式 | 说明 | 输入 | 模板 | 输出 | 内部流程 |
|------|------|------|------|------|---------|
| A: Word 格式整理 | 格式化 Word 文档 | .docx | 需要 | .docx | Word → MD → Word |
| B: Word → MD | 提取 Markdown | .docx | 不需要 | .md | Word → MD |
| C: MD → Word | 生成格式化 Word | .md | 需要 | .docx | MD → Word |

#### 操作流程

1. **选择模式** — 通过顶部单选按钮切换
2. **添加文件** — 点击"添加文件"选择单个或多个文件，或点击"添加文件夹"递归搜索目录下所有匹配文件（含子目录）。状态 C 搜索 `.md` 文件，其余状态搜索 `.docx` 文件
3. **勾选文件** — 文件列表支持逐项勾选/取消，提供"全选"和"取消全选"按钮。仅处理已勾选的文件
4. **选择模板** — 从下拉列表中选择模板（模式 B 不需要，模板区域自动隐藏）
5. **选择输出目录** — 指定处理后文件的输出位置
6. **开始处理** — 点击按钮执行，进度条和日志区域显示处理状态

#### 模板共享

1. 在模板编辑标签页中，选择目标模板，点击 **导出**
2. 将生成的 `.json` 文件发送给其他用户
3. 用户在程序中点击 **导入**，选择该 `.json` 文件，模板即加入其模板列表

---

## Markdown 格式规范

程序采用分层渐进的格式设计：所有特殊语法均为可选，用户使用标准 Markdown 即可生成格式化 Word。

### L1 — 基础语法

标准 Markdown，无需任何额外语法：

```markdown
# 一级标题
## 二级标题

正文段落，**粗体**，*斜体*，行内 `代码`。

![架构图](images/arch.png)
*图 系统架构图*

| 参数 | 说明 |
|------|------|
| 端口 | 8080 |

- 列表项1
- 列表项2

```python
def hello():
    print("world")
```

内联公式 $E=mc^2$

$$
F = ma
$$
```

系统自动完成图表编号、标题书签生成、模板字体段落套用。

### L2 — YAML Frontmatter

在 Markdown 文件头部添加可选的元数据块：

```yaml
---
cover_enabled: true
cover_fields:
  contract_no: "HT-2026-001"
  organization: "XX公司"
signature_enabled: true
revision_enabled: true
toc_enabled: true
---
```

所有字段均为可选，缺项使用模板默认值。

### L3 — 精确控制

用于 Word → MD → Word 内部管道，支持完整的 Word 语义表达：

```markdown
<!-- section: cover -->
<!-- section: body -->

## 系统设计  {#sec:design}

![架构图](images/arch.png)
*图 1 系统架构* {#fig:arch}

*表 1 参数列表* {#tbl:params}
| 参数 | 值 |
|------|-----|

$$
E = mc^2
$$ {#eq:einstein}

如图 1 所示...           → 生成 REF 域，点击可跳转到目标图表
参考：[官网](https://example.com)  → 生成 HYPERLINK 外部超链接

术语[^note1]              → 生成脚注
[^note1]: 脚注内容。

<!-- visio: attachments/diagram.vsd -->  → 嵌入 Visio OLE 对象
```

### MD 语法 → Word 元素对照

| Markdown | Word 元素 |
|----------|----------|
| `# ~ #####` | Heading 段落，套用对应级别 HeadingConfig |
| 普通段落 | 正文段落，套用 BodyConfig |
| `**text**` / `*text*` | Bold / Italic run |
| `` `code` `` | 等宽字体 run |
| ` ```lang ... ``` ` | 代码块段落 + 底纹 |
| `![alt](path)` + 题注 | 图片 + 题注段落 + Bookmark |
| `*题注* {#id}` + 表格 | 题注 + Word Table + Bookmark |
| `$...$` / `$$...$$` | LaTeX → PNG → 嵌入 |
| `[text](#fig:xxx)` | REF 域 → Bookmark（可点击跳转） |
| `[text](https://...)` | HYPERLINK |
| `[^id]` | FOOTNOTE |
| `{#id}` | Bookmark 锚点 |
| `<!-- section: xxx -->` | 分节符 |
| `<!-- visio: path -->` | Visio OLE 嵌入 |

### 降级策略

| 缺失项 | 降级行为 |
|--------|---------|
| 无 YAML frontmatter | 使用模板默认值，全文作为正文 |
| 无 `<!-- section -->` | 整个文档作为单节 |
| 无 `{#id}` 书签 | 自动按序生成 `fig:N` / `tbl:N` / `eq:N` / `sec:heading_N` |
| `[text](#id)` 目标不存在 | 渲染为纯文本 |
| 图片文件不存在 | 插入占位符文本，记录警告日志 |
| LaTeX 渲染失败 | 插入原始 LaTeX 源码及降级提示 |

---

## 常见问题

### 程序启动异常

**现象**：双击 exe 无响应或窗口一闪即退

**排查**：从命令行运行以查看错误信息：
1. 按 `Win + R`，输入 `cmd`，回车
2. 将 `DocFormatter.exe` 拖入命令行窗口
3. 按回车，查看错误输出

**高 DPI 显示异常**：右键 `DocFormatter.exe` → 属性 → 兼容性 → 更改高 DPI 设置 → 勾选"替代高 DPI 缩放行为"。

### 处理结果不符合预期

- 确认模板中"正文"字体与段落配置正确
- 如果源文件正被 Word 打开，关闭 Word 后重试
- 如果源文件已损坏，用 Word 另存为新文件后再试

### 模板管理

- 用户模板存储路径（按平台）：
  - Windows：`%APPDATA%/DocFormatter/templates/`
  - macOS：`~/Library/Application Support/DocFormatter/templates/`
  - Linux：`$XDG_DATA_HOME/DocFormatter/templates/`（默认 `~/.local/share/DocFormatter/templates/`）
- 删除对应 `.json` 文件即可移除用户模板
- 内置模板不允许删除

### 字体显示

模板编辑器中的字体列表读取自当前系统已安装字体。如果配置的字体在目标系统上不可用，格式化时将回退为宋体（SimSun）。

---

## 项目结构

```
WordFormat-Covert/
├── build.bat                        # Windows 打包脚本
├── build.sh                         # Linux/macOS 打包脚本
├── docformatter/
│   ├── main.py                      # 程序入口
│   ├── gui/
│   │   ├── main_window.py           # 主窗口
│   │   ├── template_config.py       # 模板编辑标签页
│   │   └── doc_process.py           # 文档处理标签页
│   ├── core/
│   │   ├── word2md_converter.py     # Word → Markdown 转换器
│   │   ├── md2word_converter.py     # Markdown → Word 转换器
│   │   ├── document_analyzer.py     # Word 文档结构分析
│   │   ├── header_footer.py         # 页眉页脚管理器
│   │   ├── cover_replacer.py        # 封面占位符替换
│   │   └── table_handler.py         # 表格格式化
│   ├── models/
│   │   └── template_config.py       # 数据模型定义
│   ├── templates/
│   │   ├── template_io.py           # 模板序列化
│   │   ├── template_manager.py      # 模板管理器（CRUD/导入导出）
│   │   └── builtin/
│   │       └── default.json         # 内置默认模板
│   ├── utils/
│   │   ├── font_formatter.py        # 字体/段落格式化
│   │   ├── placeholder_scanner.py   # 占位符扫描
│   │   └── logger.py               # 日志
│   └── requirements.txt
└── README.md
```
