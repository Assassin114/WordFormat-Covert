# DocFormatter 使用问题与解答

## 用户反馈的问题清单

### 1. 字号显示
**现状：** 用"pt"（磅）显示
**要求：** 改为中文标准：一号、小一、小二、三号、小三、四号、小四、五号、小五、六号、小六

**字号的对应关系：**
| 中文名称 | 磅值 |
|---------|------|
| 初号 | 42pt |
| 小初 | 36pt |
| 一号 | 26pt |
| 小一 | 24pt |
| 二号 | 22pt |
| 小二 | 18pt |
| 三号 | 16pt |
| 小三 | 15pt |
| 四号 | 14pt |
| 小四 | 12pt |
| 五号 | 10.5pt |
| 小五 | 9pt |
| 六号 | 7.5pt |
| 小六 | 6.5pt |

### 2. 字体设置中英文区分
**要求：** 中文字体和英文字体分开设置
- 中文字体（如：宋体、黑体、楷体、仿宋）
- 英文字体（如：Times New Roman、Arial、Courier New）

### 3. 字体颜色选项
**要求：** 在格式设置中增加字体颜色选择器

### 4. 字体库从主机读取
**要求：** 自动读取Windows系统的字体库，显示可用字体列表

### 5. 行距配置与Word对齐
**现状：** 只支持"倍数"行距
**要求：** 支持Word的四种行距模式：
- **固定值**：设置具体磅值
- **最小值**：确保内容不小于指定磅值
- **单倍行距**：1.0倍
- **多倍行距**：如1.5倍、2倍等

### 6. 封面模板预览
**要求：** 
- 封面模板配置增加预览功能
- 可以直接编辑和调整封面元素
- 所见即所得

### 7. 签署页固定
**要求：** 签署页（拟制/校对/标审/审核/批准）基本固定，不作为用户可配置项

### 8. 题注格式完善
**要求：** 题注（图序、表序）增加字体和段落设置选项

### 9. 打印模式单选
**要求：** 打印模式（单面打印/双面打印）改为单选按钮组，不能同时选中

### 10. 页眉页脚的作用
**说明：**
- **封面/签署页/修改记录页**：无页眉页脚
- **目录页**：大写罗马数字页码（I, II, III...）
- **正文部分**：大写罗马数字页码（从I开始，新编号）

### 11. JSON模板文件的作用
**说明：** `template.json` 是格式化配置的模板文件，包含：
- 标题1-5级的字体和段落配置
- 正文字体和段落配置
- 图序/表序/公式编号格式
- 页眉页脚设置
- 封面占位符内容

**优势：**
- 可复用：同一套配置可用于多个文档
- 可导出/导入：方便分享和备份
- 版本管理：可以看到配置的变更历史

---

## 功能优先级

### P0 - 必须修复
1. 字号改为中文名称
2. 行距配置完善（固定值/最小值/单倍/多倍）
3. 打印模式改为单选

### P1 - 应该实现
4. 字体颜色选项
5. 字体库从系统读取
6. 题注格式完善

### P2 - 可以优化
7. 字体中英文区分
8. 封面预览功能
9. 签署页简化

---

## 页眉页脚详细设计

### 页面类型分类

| 页面类型 | 内容 | 页眉 | 页脚 | 页码格式 |
|---------|------|------|------|---------|
| 第1类 | 封面页、签署页、修改记录页 | 统一文本 | 无 | 无 |
| 第2类 | 目录页 | 统一文本 | 有 | 可选 |
| 第3类 | 正文部分 | 统一文本 | 有 | 可选 |

### 页码格式（Word标准）

| 格式 | 示例 |
|------|------|
| 阿拉伯数字 | 1, 2, 3, ... |
| 小写罗马数字 | i, ii, iii, iv, ... |
| 大写罗马数字 | I, II, III, IV, ... |
| 小写字母 | a, b, c, ... |
| 大写字母 | A, B, C, ... |

### 页码位置
- 页脚（页面底部）
- 页眉（页面顶部）

### 页码对齐
- 左对齐
- 居中
- 右对齐

### 页码字体配置
- 字体名称（中文/英文）
- 字体大小
- 加粗/斜体

### GUI 配置项

```
页眉页脚设置
├── 封面页/签署页/修改记录页
│   ├── 页眉文本：[输入框]
│   ├── 显示页眉：☑/☐
│   └── 页脚：无
├── 目录页
│   ├── 页眉文本：[输入框]
│   ├── 显示页眉：☑/☐
│   ├── 显示页脚：☑/☐
│   ├── 页码格式：[下拉：阿拉伯数字/小写罗马/大写罗马/小写字母/大写字母]
│   ├── 页码位置：[下拉：页眉/页脚]
│   └── 页码对齐：[下拉：左/中/右]
├── 正文部分
│   ├── 页眉文本：[输入框]
│   ├── 显示页眉：☑/☐
│   ├── 显示页脚：☑/☐
│   ├── 页码格式：[下拉]
│   ├── 页码位置：[下拉]
│   ├── 页码对齐：[下拉]
│   ├── 页码字体：[选择框]
│   └── 页码字号：[选择框]
└── 首页不同：☑/☐
    奇偶页不同：☑/☐
```

### 数据模型设计

```python
class PageNumberConfig:
    format: str  # "arabic", "roman_low", "roman_up", "letter_low", "letter_up"
    position: str  # "header", "footer"
    alignment: str  # "left", "center", "right"
    font_name: str = "Times New Roman"
    font_size: float = 10.5
    font_bold: bool = False
    font_italic: bool = False

class HeaderFooterConfig:
    cover_header_text: str = "文档标题"
    cover_show_header: bool = True
    cover_show_footer: bool = False
    
    toc_header_text: str = "目录"
    toc_show_header: bool = True
    toc_show_footer: bool = True
    toc_page_number: PageNumberConfig
    
    body_header_text: str = "正文"
    body_show_header: bool = True
    body_show_footer: bool = True
    body_page_number: PageNumberConfig
    
    different_first_page: bool = False
    different_odd_even: bool = False
```

---

## 统一字体格式配置

### 设计原则
所有字体格式配置都通过 `FontFormatter` 统一调用，确保一致性。

### FontFormatter 工具类

```python
class FontFormatter:
    """统一字体格式配置"""
    
    @staticmethod
    def apply_font(run, font_config: FontConfig):
        """应用字体到 run"""
        run.font.name = font_config.name
        run.font.size = Pt(font_config.size / 2)  # 转换为半磅
        run.bold = font_config.bold
        run.italic = font_config.italic
        if font_config.color:
            run.font.color.rgb = RGBColor.from_string(font_config.color)
    
    @staticmethod
    def apply_paragraph(para, paragraph_config: ParagraphConfig):
        """应用段落格式"""
        # 行距
        if paragraph_config.line_spacing_type == "fixed":
            para.paragraph_format.line_spacing = paragraph_config.line_spacing_fixed
        elif paragraph_config.line_spacing_type == "atLeast":
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            para.paragraph_format.line_spacing = paragraph_config.line_spacing_min
        elif paragraph_config.line_spacing_type == "single":
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        else:  # multiple
            para.paragraph_format.line_spacing = paragraph_config.line_spacing
        
        # 段前段后
        para.paragraph_format.space_before = Pt(paragraph_config.space_before)
        para.paragraph_format.space_after = Pt(paragraph_config.space_after)
        
        # 首行缩进
        if paragraph_config.first_line_indent > 0:
            para.paragraph_format.first_line_indent = Pt(paragraph_config.first_line_indent)
```

### 表格字体配置

```python
class TableFontConfig:
    """表格字体配置"""
    # 基础表（常规表格）
    regular_table: TableStyleConfig
    
    # 记录表（测试/试验记录表）
    record_table: TableStyleConfig

@dataclass
class TableStyleConfig:
    """表格样式配置"""
    header_font: FontConfig      # 表头字体
    header_bg_color: str        # 表头背景色
    body_font: FontConfig       # 正文字体
    header_align: str            # 表头对齐
    body_align_rule: str         # 正文对齐规则（按字数）
    body_align_threshold: int    # 短文本阈值
```

### 使用位置
- 标题格式化 → FontFormatter
- 正文格式化 → FontFormatter
- 表格格式化 → FontFormatter + TableFontConfig
- 题注格式化 → FontFormatter
- 页眉页脚 → FontFormatter
- 页码 → FontFormatter

---

## 行距配置（GUI二级菜单设计）

### 当前行距设置方式
四个选项通过下拉框选择，选中后显示对应的值设置：

```
行距类型: [单倍行距 ▼]
            ├─ 固定值 → 显示"磅值(P): [20] pt
            ├─ 最小值 → 显示"磅值(P): [15] pt  
            ├─ 单倍行距 → 显示"单倍行距"(不可改)
            └─ 多倍行距 → 显示"倍数(M): [1.5] x
```

### 段前/段后/首行缩进
```
段落格式:
├─ 段前间距: [0] pt
├─ 段后间距: [0] pt
└─ 首行缩进: [2] 字符
```

---

## GUI布局设计（左右分栏）

### 整体布局
```
+--------------------------------------------------+
|  [新建] [加载模板] [保存模板]                      |
+--------------------------------------------------+
|         |                                          |
|  项目列表  |         配置面板                        |
|  --------  |  ---------------------                 |
|  标题1    |  字体: [宋体 ▼] [Times New Roman ▼]    |
|  标题2    |  字号: [五号 ▼]                       |
|  标题3    |  颜色: [■] 加粗: [■] 斜体: [ ]        |
|  标题4    |  ----                                   |
|  标题5    |  行距类型: [多倍行距 ▼]                 |
|  正文      |  倍数(M): [1.5] x                      |
|  表格-基础 |  段前: [0] pt  段后: [0] pt          |
|  表格-记录 |  首行缩进: [2] 字符                   |
|  题注-图  |                                         |
|  题注-表  |                                         |
|  题注-公式|                                         |
|  页眉-封面|                                         |
|  页眉-目录|                                         |
|  页眉-正文|                                         |
|  封面字段|                                         |
|         |                                          |
+--------------------------------------------------+
```

### 左侧项目列表
使用 QListWidget 显示所有可配置项，点击后在右侧显示对应配置

### 右侧配置面板
使用 QStackedWidget 根据左侧选择切换不同配置页面
