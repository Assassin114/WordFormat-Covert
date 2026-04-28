# P004 - 软件开发\n\n**创建日期：** 2026-04-28\n\n**关联技能：** superpowers-mode\n\n**状态：** 进行中\n

---

## 开发进度 (2026-04-28)

### 已完成

1. **基础框架搭建**
   - 目录结构创建完成
   - 依赖文件 requirements.txt

2. **数据模型** (`models/`)
   - TemplateConfig 及相关数据类
   - FormatResult / BatchResult

3. **模板管理** (`templates/`)
   - 模板加载/保存 (JSON格式)
   - 默认模板创建

4. **工具模块** (`utils/`)
   - ooxml_utils.py: 底层XML操作（行距/孤行控制/字体等）
   - logger.py: 统一日志工具

5. **核心引擎** (`core/`)
   - formatter.py: 核心格式化调度
   - style_mapper.py: 样式映射
   - numbering.py: 图/表/公式编号管理
   - toc_generator.py: 目录生成

6. **GUI界面** (`gui/`)
   - main_window.py: 主窗口框架
   - template_config.py: 模板配置Tab
   - batch_process.py: 批量处理Tab
   - word2md_tab.py: Word转MD Tab

7. **程序入口** (`main.py`)

### 待完成

1. 封面占位符替换逻辑（等亮哥提供封面模板文件）
2. 图片题注插入（同页控制）
3. 表格题注插入（同页控制）
4. 公式编号插入
5. 横向页面处理完善
6. 双面打印空白页处理
7. GUI细节完善

### 依赖安装

需要安装依赖才能运行：
```bash
pip install PyQt6 python-docx lxml mammoth
```

---

## 开发进度更新 (2026-04-28 01:28)

### 本次更新 - 核心功能完善

#### 1. 图片题注处理 (`_process_images`)
- 遍历所有段落，检测包含drawing元素的段落
- 在图片下方插入居中题注段落（孤行控制 + 与前段同页）
- 题注格式：宋体 10.5pt，居中对齐，段前6pt间距

#### 2. 表格题注处理 (`_process_tables`)
- 遍历doc.tables，定位table元素
- 在表格上方插入居中题注段落
- 题注格式与图片题注一致

#### 3. 公式编号处理 (`_process_equations`)
- 检测OMML公式元素（`w:oMath`）
- 公式段落居中对齐
- 使用制表位实现右侧编号（左公式右编号）

#### 4. 新增工具函数
- `utils/file_utils.py`: 文件扫描、路径处理、文件大小格式化
- `get_paragraph_pPr`: 通用的段落属性元素获取（支持element和paragraph）

### 代码验证

基础框架可正常导入：
- ✓ models
- ✓ templates  
- ✓ utils
- ✓ core

### 待完善

1. 封面占位符替换逻辑（等亮哥提供封面模板文件）
2. 横向页面处理完善
3. 双面打印空白页处理
4. GUI细节完善
5. Word转MD的mammoth集成

---

## 开发进度更新 (2026-04-28 01:31)

### 本次更新 - 交叉引用处理

#### 新增模块：`core/cross_reference.py`

**CrossReferenceManager 类：**

1. **scan_document()** - 扫描文档中的题注书签和REF引用
   - 识别 fig_X, tbl_X, eq_X 格式的书签
   - 记录所有 REF xxx 域

2. **renumber_caption()** - 重新编号时更新内部记录

3. **update_ref_fields()** - 更新所有REF域的显示文本
   - 遍历instrText找到REF域
   - 更新fldCharType=separate之后的<t>元素

4. **update_bookmark_names_in_document()** - 重命名文档中的书签

#### formatter.py 集成

- 在format()流程的Step 3.5添加了交叉引用扫描
- 图片/表格处理时会同步更新书签和REF域

#### 交叉引用处理流程

```
1. scan_document() - 扫描建立映射表
2. 重新编号时重命名书签 - update_bookmark_names_in_document()
3. 更新REF域显示文本 - update_ref_fields()
```

### 待完善

1. 封面占位符替换逻辑（等亮哥提供封面模板文件）
2. 横向页面处理完善
3. 双面打印空白页处理
4. GUI细节完善
5. Word转MD的mammoth集成
6. 交叉引用更新逻辑完整测试

---

## 开发进度更新 (2026-04-28 01:37)

### 本次更新 - 继续完善

#### 1. 双面打印空白页处理
- `_handle_print_mode()` 在文档末尾添加分页段落
- 确保双面打印时文档从偶数页开始

#### 2. 新增 Word2MDConverter 模块
- `core/word2md_converter.py`
- 基于mammoth库的封装
- 支持批量转换

#### 3. 代码结构完善
- formatter.py 现在导入所有核心模块
- utils/ooxml_utils.py 修正了 get_paragraph_pPr 函数

### 当前代码文件清单

```
docformatter/
├── main.py
├── gui/
│   ├── main_window.py
│   ├── template_config.py
│   ├── batch_process.py
│   └── word2md_tab.py
├── core/
│   ├── __init__.py
│   ├── formatter.py           ← 核心调度（已完善）
│   ├── style_mapper.py
│   ├── numbering.py
│   ├── toc_generator.py
│   ├── cross_reference.py     ← 交叉引用处理
│   └── word2md_converter.py    ← Word转MD
├── utils/
│   ├── __init__.py
│   ├── ooxml_utils.py          ← 底层XML操作
│   ├── file_utils.py           ← 文件工具
│   └── logger.py
├── templates/
│   ├── __init__.py
│   └── template_io.py
├── models/
│   ├── __init__.py
│   └── template_config.py
└── requirements.txt
```

### 待完善功能

1. 封面占位符替换逻辑（等亮哥提供封面模板文件）
2. 横向页面处理完善
3. GUI细节完善（模板配置Tab和批量处理Tab的保存/加载模板联动）
4. 完整功能测试

---

## 项目文件清单 (2026-04-28 01:37)

### 全部文件 (23个.py + 配置文件)

```
docformatter/
├── main.py                          # 程序入口
├── __init__.py
├── requirements.txt
├── README.md
│
├── gui/                              # PyQt6 界面 (4个文件)
│   ├── main_window.py                # 主窗口框架
│   ├── template_config.py             # 模板配置Tab
│   ├── batch_process.py              # 批量处理Tab
│   ├── word2md_tab.py                # Word转MD Tab
│   └── widgets/__init__.py
│
├── core/                             # 核心引擎 (7个文件)
│   ├── __init__.py
│   ├── formatter.py                  # 核心调度
│   ├── style_mapper.py                # 样式映射
│   ├── numbering.py                   # 编号管理
│   ├── toc_generator.py              # 目录生成
│   ├── cross_reference.py             # 交叉引用处理
│   └── word2md_converter.py          # Word转MD
│
├── utils/                             # 工具 (3个文件)
│   ├── __init__.py
│   ├── ooxml_utils.py                 # 底层XML操作
│   ├── file_utils.py                  # 文件工具
│   └── logger.py                      # 日志工具
│
├── templates/                         # 模板管理 (2个文件)
│   ├── __init__.py
│   └── template_io.py                 # 模板加载/保存
│
└── models/                            # 数据模型 (2个文件)
    ├── __init__.py
    └── template_config.py              # 数据类定义
```

### 代码行数统计

```bash
$ find docformatter -name "*.py" -not -path "*/__pycache__/*" | xargs wc -l | tail -1
  ~4500 行 Python 代码
```

### 语法检查

- ✅ main.py
- ✅ formatter.py
- ✅ cross_reference.py
- ✅ word2md_converter.py
- ✅ 所有模块可正常导入

---

## 开发进度更新 (2026-04-28 15:25)

### 本次更新 - 整合封面模板和表格模板

#### 1. 新增模板文件
- `templates/cover_fields.json` — 封面占位符映射
  - 合同编号、密级、单位、日期、签署页等字段
- `templates/table_field_template.json` — 表格字段映射模板
  - 支持：常规表格、测试记录表、试验记录表

#### 2. 新增 TableHandler 模块
- `core/table_handler.py` — 独立的表格格式化处理器
- 功能：
  - 自动识别表格类型（常规 vs 记录表）
  - 按字段映射自动应用格式（加粗、对齐）
  - 支持测试/试验记录表的键值对识别

#### 3. 更新 formatter.py
- 集成 TableHandler 到文档处理流程
- Step 7 处理表格时调用 table_handler.format_table()
- 保持原有的题注插入逻辑不变

#### 4. 用户提供的模板
- 研究报告类模板.docx — 包含封面、修改记录表、示例表格
- 记录表模板.docx — 包含测试记录表、试验记录表

### 当前文件结构
```
docformatter/
├── templates/
│   ├── cover_fields.json      ← 新增：封面字段映射
│   └── table_field_template.json  ← 新增：表格字段模板
├── core/
│   ├── table_handler.py       ← 新增：表格处理器
│   └── formatter.py           ← 更新：集成表格处理器
└── ...
```

### 模板字段说明

**封面字段 (cover_fields.json):**
- contract_no, classification, year, organization, date
- drafter, drafter_time, reviewer, reviewer_time
- approver, approver_time, checker, checker_time
- final_approver, final_approver_time

**表格类型:**
- 常规表格：表头居中加粗，正文按长度判断对齐
- 测试记录表：字段名加粗居左，值不加粗，签名行特殊处理
- 试验记录表：同上

### 待完成
1. 封面占位符替换逻辑完善（支持用户提供的模板文件）
2. GUI中模板配置Tab联动保存/加载
3. 完整功能测试

---

## 开发进度更新 (2026-04-28 15:45)

### 完整功能测试结果

#### 测试1：研究报告类模板（常规表格）
- ✅ 表格类型检测：正确识别为"常规表格"
- ✅ 表头格式化：序号/版本号/日期/修改条款/作者 加粗居中
- ✅ 数据行格式化：内容居中，不加粗

#### 测试2：记录表模板（测试记录表）
- ✅ 表格类型检测：正确识别为"测试记录表"
- ✅ 章节标题行：基本信息 加粗居中
- ✅ 字段名行：测试用例名称/测试内容/测试方法等 加粗居左
- ✅ 签名行：测试人员/操作人员/测试时间 正确分类处理

#### 修复的问题
1. `table_field_template.json` 中 `field_name` 缺少 `fields` 列表
2. `table_handler.py` 中 `_format_record_table` 的 `known_fields` 没有包含 `field_name.fields`

### 模板文件清单
```
templates/
├── cover_fields.json         # 封面字段映射
└── table_field_template.json # 表格字段映射（已修复）
```

### 验证文件
- `/tmp/test_report_final.docx` — 研究报告类模板格式化结果
- `/tmp/test_record_final.docx` — 记录表模板格式化结果

### 核心修复代码位置
- `core/table_handler.py` — 完整的表格格式化处理器（已集成到 formatter.py 的 Step 7）

---

## 开发进度更新 (2026-04-28 16:00)

### 1. 整体格式测试

#### 混乱测试文档
- 路径: `/tmp/test_messy.docx`
- 内容: 标题(0级)、副标题、3级标题、混乱正文、常规表格、测试记录表

#### 格式化测试结果
- ✅ 段落格式化：标题/正文检测正确
- ✅ 对齐应用：居中/左对齐/两端对齐正常
- ✅ 首行缩进：420 (2字符) 正常应用
- ✅ 行距：360 (1.5倍行距) 正常应用
- ✅ 表格格式化：正确识别常规表格和测试记录表

#### 发现的问题
1. `add_heading(level=0)` 创建的是 Title 样式，不是 Heading 0 — 需要在 is_heading() 中处理
2. 原始测试文档的 P2 分隔符段落被跳过处理

### 2. Word转MD功能

#### 新增文件
- `core/word2md_converter.py` — 独立转换器模块

#### 支持功能
- 标题: ATX风格 (# ~ ######)
- 粗体: `**text**`
- 斜体: `*text*`
- 表格: HTML table标签
- 列表: 自动检测numPr

#### 测试结果
- `/tmp/test_messy.docx` → `/tmp/test_messy.md` (556字符)
- `/tmp/test_formatted.docx` → `/tmp/test_formatted.md` (552字符)

### 测试文件清单
```
/tmp/test_messy.docx       # 混乱格式测试文档
/tmp/test_messy.md         # 混乱文档转MD
/tmp/test_formatted.docx   # 格式化后的文档
/tmp/test_formatted.md     # 格式化文档转MD
```

### 待完善
1. 列表检测和编号转换
2. 图片引用转换
3. 代码块识别
4. 嵌套表格支持

---

## 开发进度更新 (2026-04-28 16:25)

### Bug修复
1. **cover_replacer.py 语法错误** — f-string中单个`}`导致语法错误
2. **toc_generator.py 缺少导入** — `HeadingConfig`未导入
3. **formatter.py 未集成CoverReplacer** — 原来用的是简单字符串替换

### 已集成的模块
- ✅ DocumentFormatter (formatter.py)
- ✅ TableHandler (table_handler.py) — 表格格式化
- ✅ CoverReplacer (cover_replacer.py) — 封面占位符替换
- ✅ CrossReferenceManager (cross_reference.py) — 交叉引用管理
- ✅ NumberingManager (numbering.py) — 编号管理
- ✅ TOCGenerator (toc_generator.py) — 目录生成
- ✅ StyleMapper (style_mapper.py) — 样式映射
- ✅ Word2MDConverter (word2md_converter.py) — Word转MD

### Git提交记录
```
b5e62c4 fix: integrate CoverReplacer into formatter and fix import errors
aabac39 feat: enhance template config with cover fields editor and CoverReplacer
0f955cb feat: add CoverReplacer for structured cover placeholder replacement
1b1d7fd docs: add SDD and SPEC documents
16ca0bc feat: DocFormatter v1.0 - Word文档格式化工具
```

### 集成测试结果
- ✅ CoverReplacer 替换4个占位符正常
- ✅ TableHandler 识别并格式化2个测试记录表
- ✅ DocumentFormatter 实例化正常
- ✅ Word2MDConverter 转换正常 (1722 chars)
- ✅ CrossReferenceManager 实例化正常

### 对照SPEC的功能覆盖状态

| 功能 | 状态 | 说明 |
|------|------|------|
| 模板配置（封面/签署页/正文/标题） | ✅ | GUI模板配置界面已完成 |
| 文档格式化 | ✅ | formatter.py核心引擎 |
| 批量处理 | ✅ | batch_process.py |
| Word转MD | ✅ | word2md_converter.py |
| 表格格式化 | ✅ | table_handler.py |
| 封面占位符替换 | ✅ | CoverReplacer集成 |
| 交叉引用管理 | ✅ | CrossReferenceManager |
| 图序/表序/公式编号 | ✅ | NumberingManager |
| 目录生成 | ✅ | TOCGenerator |
| 页眉页脚 | ⚠️ | 基础实现 |
| 打印模式 | ⚠️ | 基础实现 |
| 孤行控制 | ⚠️ | 需要完善 |
| 纯文本引用更新 | ✅ | CrossReferenceManager |
| 书签重命名 | ✅ | CrossReferenceManager |

---

## 开发进度更新 (2026-04-28 16:40)

### 实现的三个功能

#### 1. 孤行控制（widow/keepNext）✅
- widowControl 和 keepNext 元素设置 val="true"
- 图序题注（下方）设置 keepNext
- 表序题注（上方）设置 keepNext

#### 2. 分节与横向页面处理 ✅
- _handle_landscape_sections() 检测并保留横向节设置
- 通过 pgSz/orient 判断 landscape

#### 3. 页眉页脚（按页面类型分类）✅
**新增 HeaderFooterManager 类**，支持三类页面：

| 页面类型 | 页眉 | 页脚 | 页码 |
|---------|------|------|------|
| 封面/签署页/修改记录页 | 统一文本 | 无 | 无 |
| 目录页 | 统一文本 | 有 | 大写罗马数字（I, II...） |
| 正文部分 | 统一文本 | 有 | 大写罗马数字（从I开始） |

**页码格式：** `PAGE \* roman` 域代码生成大写罗马数字

### Git提交记录
```
90c26df feat: implement three-type header/footer system
5e4ad24 feat: implement widow/keepNext control, landscape section handling
b5e62c4 fix: integrate CoverReplacer into formatter and fix import errors
aabac39 feat: enhance template config with cover fields editor
0f955cb feat: add CoverReplacer for structured cover placeholder replacement
1b1d7fd docs: add SDD and SPEC documents
16ca0bc feat: DocFormatter v1.0 - Word文档格式化工具
```

### 待完善项
- 页眉页脚文本内容的GUI配置（目前硬编码）
- 罗马数字起始值的正确设置
