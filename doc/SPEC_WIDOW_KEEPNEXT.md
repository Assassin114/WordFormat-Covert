# 孤行控制与分节处理规格说明

## 1. 孤行控制（Widow/Keep-Next）

### 1.1 功能描述

| 元素 | 位置 | 同页要求 |
|------|------|---------|
| 图序 | 图下方 | 与上图同页（keepNext） |
| 表序 | 表上方 | 与下表同页（keepNext） |
| 公式 | 居中 | 无特殊要求 |

### 1.2 技术实现

```python
# 孤行控制 - widowControl
# w:widowControl val="true" 启用

# 同页控制 - keepNext  
# w:keepNext val="true" 确保与下一段同页
```

### 1.3 优先级
- **P0**：图序/表序的 keepNext 控制（核心）
- **P1**：公式编号居中对齐

---

## 2. 分节与横向页面处理

### 2.1 功能描述

- 保留分节后的横向页面设置
- 横向节前后有分节符
- 横向节内可能有1~N张图片
- 后续节自动恢复纵向

### 2.2 技术实现

```python
# 检测横向页面
section = doc.sections[0]
pgSz = section._element.find(qn('w:pgSz'))
orient = pgSz.get(qn('w:orient'))  # 'portrait' or 'landscape'

# 判断横向
is_landscape = orient == 'landscape' or (w and h and int(w) > int(h))

# 横向节：保留页面尺寸和方向设置，不插入额外分页符
```

### 2.3 处理流程

```
1. 遍历所有 section
2. 检查 pgSz/orient 属性
3. 如果是 landscape，标记为横向节
4. 后续处理时不在横向节内插入额外分页符
5. 横向节后的下一节自动恢复纵向
```

---

## 3. 页眉页脚设置

### 3.1 功能描述

- 首页不同（different_first_page_header_footer）
- 奇偶页不同（even_odd_header_footer）
- 页码格式设置

### 3.2 当前实现

`formatter.py` 的 `_apply_header_footer` 方法已有基础实现：

```python
def _apply_header_footer(self, doc: Document):
    for section in doc.sections:
        if self.template.header_footer.show_on_first:
            section.different_first_page_header_footer = True
        if self.template.header_footer.different_odd_even:
            section.even_odd_header_footer = True
```

### 3.3 待完善

- 页眉内容设置（文字/页码）
- 页脚内容设置
- 首页/奇偶页不同内容

---

## 4. 验收标准

1. ✅ 图序段落设置 `keepNext`，确保与图同页
2. ✅ 表序段落设置 `keepNext`，确保与表同页
3. ✅ 公式居中编号
4. ✅ 横向节分节符保留
5. ✅ 横向节内图片保留
6. ✅ 横向节后恢复纵向
7. ✅ 首页不同/奇偶页不同的设置生效
