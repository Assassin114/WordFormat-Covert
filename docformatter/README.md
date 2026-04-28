# DocFormatter

Word文档格式批量处理工具

## 功能

- **模板配置**: 通过GUI配置文档格式模板（封面/签署页/标题/正文/题注等）
- **批量处理**: 批量选择多个Word文档，使用同一模板进行格式化
- **自动编号**: 图序/表序/公式序全文统一编号
- **目录生成**: 支持自动生成目录
- **分节处理**: 支持分节后的横向页面
- **Word转MD**: 将Word文档转换为Markdown

## 技术栈

- Python 3.11
- PyQt6 (GUI)
- python-docx (Word读写)
- lxml (底层XML操作)
- mammoth (Word转Markdown)

## 安装

```bash
# 1. 创建虚拟环境（推荐）
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 2. 安装依赖
pip install PyQt6 python-docx lxml mammoth

# 或使用项目中的 requirements.txt
pip install -r requirements.txt
```

## 运行

```bash
python main.py
```

## 项目结构

```
docformatter/
├── main.py                     # 程序入口
├── gui/                        # PyQt6 GUI模块
│   ├── main_window.py          # 主窗口
│   ├── template_config.py      # 模板配置界面
│   ├── batch_process.py        # 批量处理界面
│   └── word2md_tab.py          # Word转MD界面
├── core/                       # 核心引擎
│   ├── formatter.py            # 核心格式化引擎
│   ├── style_mapper.py         # 样式映射
│   ├── numbering.py           # 编号管理
│   └── toc_generator.py       # 目录生成
├── utils/                      # 工具模块
│   ├── ooxml_utils.py          # OOXML底层操作
│   └── logger.py               # 日志工具
├── templates/                  # 模板
│   └── template_io.py         # 模板加载/保存
├── models/                     # 数据模型
│   └── template_config.py      # 模板配置数据类
└── requirements.txt
```

## 使用流程

### 1. 配置模板

1. 启动程序，进入"模板配置"Tab
2. 配置各级标题格式、正文格式、题注格式等
3. 选择封面模板文件（可选）
4. 保存模板为JSON文件

### 2. 批量处理

1. 进入"批量处理"Tab
2. 选择要处理的Word文档（或选择整个文件夹）
3. 加载保存的模板文件
4. 选择输出目录
5. 点击"开始处理"

### 3. Word转MD

1. 进入"Word转MD"Tab
2. 选择要转换的Word文档
3. 点击"转换为Markdown"

## 注意事项

- 程序需要Windows环境运行（Python 3.11 + PyQt6）
- 处理大量文档时可能需要较长时间，请耐心等待
- 批量处理时，单个文件失败不会影响其他文件
- 封面模板需要使用 `{{FIELD_NAME}}` 格式的占位符

## 状态

**开发中** - 基础框架已完成，等待亮哥提供封面模板文件后进一步完善
