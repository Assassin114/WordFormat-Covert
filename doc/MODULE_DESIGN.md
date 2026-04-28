# DocFormatter - 功能与模块拆解

## 一、系统架构图

```
┌──────────────────────────────────────────────────────────────┐
│                        DocFormatter                           │
├──────────────────────────────────────────────────────────────┤
│  ┌──────────────────┐  ┌──────────────────┐                  │
│  │   GUI Layer       │  │   Core Engine    │                  │
│  │  (PyQt6)          │──│  (Pure Python)   │                  │
│  └──────────────────┘  └──────────────────┘                  │
│           │                     │                            │
│           ▼                     ▼                            │
│  ┌──────────────────────────────────────────┐                │
│  │            模块组成                         │                │
│  └──────────────────────────────────────────┘                │
└──────────────────────────────────────────────────────────────┘
```

## 二、GUI模块（gui/）

### 2.1 主窗口模块（main_window.py）
**职责:** 程序入口，Tab导航，菜单栏

**功能清单:**
- 窗口初始化（800x600，最小尺寸）
- Tab切换：模板配置 / 批量处理 / Word转MD
- 菜单：文件（新建/打开/保存模板，退出）、帮助（关于、使用说明）
- 状态栏显示处理进度

**对外接口:**
```python
class MainWindow(QMainWindow):
    def load_template(self, path: str)
    def save_template(self, path: str)
    def get_current_template() -> dict
    def set_template_config(self, config: dict)
```

### 2.2 模板配置模块（template_config.py）
**职责:** 模板各项配置的GUI编辑

**子组件:**

| 组件 | 功能 |
|------|------|
| CoverConfigWidget | 封面字段配置（标题/单位/日期等占位符） |
| SignatureConfigWidget | 签署页字段配置 |
| RevisionConfigWidget | 修改记录章节格式 |
| HeadingConfigWidget | 1-5级标题格式配置（字体/字号/加粗/颜色/行距/段前段后） |
| BodyConfigWidget | 正文格式配置（字体/字号/行距/首行缩进/对齐） |
| CaptionConfigWidget | 题注格式配置（标签前缀、编号格式、位置） |
| HeaderFooterConfigWidget | 页眉页脚配置（首页不同/奇偶页不同） |
| PrintModeConfigWidget | 打印模式（单面/双面） |
| TemplateSaveLoadWidget | 模板保存/加载按钮 |

**配置数据模型:**
```python
@dataclass
class TemplateConfig:
    cover: CoverConfig           # 封面字段
    signature: SignatureConfig   # 签署页字段
    revision: RevisionConfig     # 修改记录格式
    headings: List[HeadingConfig, 5]  # 5级标题
    body: BodyConfig             # 正文格式
    caption: CaptionConfig       # 题注格式
    header_footer: HeaderFooterConfig
    print_mode: PrintMode        # SINGLE/DUPLEX
```

### 2.3 批量处理模块（batch_process.py）
**职责:** 批量选择文件、启动处理、显示进度

**功能清单:**
- 文件选择（单个/文件夹/多选）
- 模板文件选择
- 输出目录选择
- 开始/取消按钮
- 进度条（QProgressBar）
- 日志显示区（QTextEdit，实时刷新）
- 处理结果统计（成功/失败数量）

### 2.4 Word转MD模块（word2md.py）
**职责:** Word转Markdown的独立功能入口

**功能清单:**
- 文件选择
- 输出路径设置
- 转换按钮
- 预览按钮（用系统默认编辑器打开生成的MD）

## 三、核心引擎模块（core/）

### 3.1 文档格式化引擎（formatter.py）
**职责:** 核心格式化逻辑，调度各子模块

**public API:**
```python
class DocumentFormatter:
    def format(self, input_path: str, output_path: str, template: TemplateConfig, cover_template: str) -> FormatResult
    def batch_format(self, input_paths: List[str], output_dir: str, template: TemplateConfig, cover_template: str) -> BatchResult
```

**内部流程:**
```
format(input_path)
  ├── parse_document()        # 解析原始文档
  ├── map_styles()            # 样式映射
  ├── apply_heading_formats() # 应用1-5级标题格式
  ├── apply_body_format()     # 应用正文格式
  ├── process_images()        # 处理图片：添加图序、同页控制
  ├── process_tables()        # 处理表格：添加表序、同页控制
  ├── process_equations()     # 处理公式：编号、居中
  ├── apply_header_footer()   # 应用页眉页脚
  ├── generate_toc()          # 生成目录
  ├── handle_print_mode()     # 处理单双面空白页
  └── save(output_path)       # 保存
```

### 3.2 样式映射模块（style_mapper.py）
**职责:** 将原始文档样式映射到目标模板

**功能:**
- 扫描原始文档所有样式（Heading1~5, Normal, Caption等）
- 建立映射规则：原始样式名 → 模板配置
- 处理直接格式（override）vs 样式应用的优先级
- 输出映射报告

**映射规则:**
```
Heading1 → 模板 Heading[0] (1级标题)
Heading2 → 模板 Heading[1] (2级标题)
...
Normal → 模板 Body
Caption → 模板 Caption
```

### 3.3 编号管理模块（numbering.py）
**职责:** 图序、表序、公式序的统一编号

**规则:**
- 全局统一序号（全文连续编号）
- 格式：图 1, 图 2, 表 1, 表 2, 公式 (1)
- 扫描文档时收集所有图片/表格/公式，分配序号
- 插入/更新题注

**public API:**
```python
class NumberingManager:
    def process(self, document) -> None
    def get_next_figure_num() -> int
    def get_next_table_num() -> int
    def get_next_equation_num() -> int
```

### 3.4 目录生成模块（toc_generator.py）
**职责:** 生成/更新Word目录

**前置条件:** 文档已完成标题格式化

**生成步骤:**
1. 收集所有Heading1~5的文本和级别
2. 构建目录结构
3. 在"目录"标题后插入域（TOC field）
4. 设置目录更新提示

**顺序强制:** 封面→签署页→修改记录→目录→正文

### 3.5 Word转MD转换器（word2md_converter.py）
**职责:** Word文档转Markdown

**使用已安装的markitdown库**

**转换规则:**
| Word元素 | MD格式 |
|----------|--------|
| Heading 1~6 | # ~ ###### |
| Bold | **text** |
| Italic | *text* |
| Table | Markdown表格 |
| Image | 保存到media/，引用路径 |
| Link | [text](url) |
| List | - item |
| Code block | ``` |

## 四、模板相关（templates/）

### 4.1 内置默认模板
**文件:** `default_template.json`
提供一套默认格式配置，用户可基于此修改

### 4.2 模板加载/保存（template_io.py）
```python
def load_template(path: str) -> TemplateConfig
def save_template(config: TemplateConfig, path: str)
```

## 五、工具模块（utils/）

### 5.1 OOXML工具（ooxml_utils.py）
**职责:** 直接操作Word底层XML，处理python-docx难以实现的格式

**场景:**
- 精确行距控制（linePitch 240）
- 段前段后距（spaceBefore/spaceAfter）
- 孤行控制（widowControl）
- 多级列表（numbering）

### 5.2 文件工具（file_utils.py）
**职责:** 文件路径处理、批量文件扫描

### 5.3 日志工具（logger.py）
**职责:** 统一日志输出到GUI和文件

## 六、数据流

```
用户配置模板
    ↓
保存为JSON (template.json)
    ↓
批量处理: 选择多个docx + template.json
    ↓
DocumentFormatter.format() 逐个处理
    ↓
输出格式化后的docx + 处理报告.log
```

## 七、模块依赖关系

```
gui/
├── main_window.py        # 依赖所有gui子模块
├── template_config.py    # 无跨层依赖
├── batch_process.py     # 依赖 core/formatter.py
└── word2md.py           # 依赖 core/word2md_converter.py

core/
├── formatter.py          # 依赖所有core子模块
├── style_mapper.py       # 无深度依赖
├── numbering.py          # 无深度依赖
├── toc_generator.py      # 依赖 numbering.py
└── word2md_converter.py  # 依赖 utils/logger.py

utils/
├── ooxml_utils.py        # 无跨层依赖
├── file_utils.py         # 无跨层依赖
└── logger.py             # 无跨层依赖
```

## 八、讨论点

1. **GUI框架:** PyQt6 vs PyQt5（兼容性） vs Tkinter（轻量）？
2. **封面模板替换机制:** 亮哥提供封面docx后，如何标记需要替换的字段？（占位符如`{{TITLE}}`？）
3. **错误处理策略:** 批量处理中单个文件失败是跳过还是中断？
4. **进度报告粒度:** 每个文件处理到哪一步（解析/格式化/保存）？
5. **Word转MD优先级:** 与格式化功能并行开发还是先做格式化？
