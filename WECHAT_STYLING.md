# 微信公众号文章 HTML 排版规范

> 配合 wechat-oa Skill 使用。解决微信草稿箱段落间距消失、标题样式不生效等问题。
> 适用版本：wechat-oa v1.x 及以上

---

## 一、核心问题

wechat-oa 推送流程中，有一个 `_clean_wechat_margins()` 函数，它会**删除所有块级元素（`<p>`、`<h1>`—`<h6>`、`<div>`）的 `margin` 属性**。

原因是微信渲染器不折叠相邻元素边距，导致 `margin` 在手机上变成两倍空白。

**结果：**
```css
p { margin: 0 0 30px 0; }   /* 被删掉了 */
h2 { margin-top: 40px; }    /* 被删掉了 */
```

段落之间完全没间距，所有内容粘在一起。

---

## 二、核心理念

**段落间距不用 `margin`/`padding`，用 `<br>` 实体换行。**

`<br>` 标签不在 `_clean_wechat_margins()` 的处理范围内，它会原样保留。

---

## 三、排版规范

### 3.1 段落间距

**✅ 正确做法：**

在每段后面插入一个 `<br>`：

```html
<p>第一段的内容...</p>
<br>
<p>第二段的内容...</p>
<br>
<h2>一级标题</h2>
<br>
<p>第三段的内容...</p>
```

不要写成两个 `<br>`（空行太多）：
```html
<p>第一段</p>
<br><br>      ← ❌ 两行太多
<p>第二段</p>
```

### 3.2 一级标题（`<h2>`）

**✅ 正确做法：**

```css
h2 {
    font-size: 17px;      /* 比正文大 2px */
    font-weight: bold;
    text-align: center;   /* 居中 */
    color: #2563eb;       /* 鲜艳蓝色 */
    line-height: 1.4;
}
```

前面插一个 `<br>`：
```html
<br>
<h2>一、核心观点</h2>
```

### 3.3 二级标题（`<h3>`）

**✅ 正确做法：**

```css
h3 {
    font-size: 15px;
    font-weight: bold;
    color: #b91c1c;       /* 鲜红色，与蓝色互补 */
    line-height: 1.4;
}
```

**⚠️ 不要有重复的 CSS 属性**（后面的覆盖前面的）：
```css
h3 { color: #b91c1c; ... color: #333; }   ← ❌ #333 会覆盖 #b91c1c
```

### 3.4 正文（`<p>`）

**✅ 正确做法：**

```css
p {
    font-size: 15px;
    color: #333;
}
```

不需要 `margin` 或 `padding`，间距用 `<br>` 控制。

### 3.5 表格

表格前面不需要 `<br>`，但后面需要：

```html
<p>下面是对比结果：</p>
<table>
    <tr><td>...</td></tr>
</table>
<br>
<h3>ChatGPT 详细分析</h3>
```

### 3.6 列表

列表内部**不要**插 `<br>`：

```html
<ul>
    <li>第一条</li>
    <li>第二条</li>
    <li>第三条</li>
</ul>
<br>
<p>列表后面的内容...</p>
```

### 3.7 引用框和提示框

`<blockquote>`、`<div class="warning-box">`、`<div class="tip-box">` 内部**不要**插 `<br>`：

```html
<div class="warning-box">
    <p>⚠️ 核心认知：...</p>
</div>
<br>
<h2>下一节标题</h2>
```

### 3.8 页脚（footer）

```css
.footer {
    text-align: center;
    color: #999;
    font-size: 14px;
    padding-top: 30px;
    border-top: 1px solid #eee;
}
```

前面插一个 `<br>`：
```html
<br>
<div class="footer">本文排版由 wechat-oa Skill 提供</div>
```

⚠️ 注意：`.footer` 的 `margin-top` 会被清理，所以用 `<br>` 来分段。`padding-top` 和 `border-top` 不受影响。

---

## 四、配色方案

| 用途 | 颜色 | 色值 | 备注 |
|------|------|------|------|
| 正文 | 深灰 | `#333` | 阅读舒适 |
| 一级标题 | 蓝色 | `#2563eb` | 居中，醒目 |
| 二级标题 | 红色 | `#b91c1c` | 与蓝色互补 |
| 表格表头 | 蓝色 | `#3498db` 背景 + `#fff` 文字 | |
| 强调文本 | 灰色 | `#333`（加粗） | 与正文同色 |
| 链接 | 蓝色 | `#1a73e8` | |
| 页脚 | 浅灰 | `#999` | 居中 |
| 警告框背景 | 淡黄 | `#fff3cd` | `border-left: #ff9800` |
| 提示框背景 | 淡绿 | `#e8f5e9` | `border-left: #4caf50` |
| 引用框背景 | 淡粉 | `#fff5f5` | `border-left: #3498db` |

**配色原则：** 不超过5种主体色，保持统一。主色用蓝色系，强调色用红色/橙色。

---

## 五、完整示例

```html
<style>
h2 { font-size: 17px; font-weight: bold; text-align: center; color: #2563eb; line-height: 1.4; }
h3 { font-size: 15px; font-weight: bold; color: #b91c1c; line-height: 1.4; }
p { font-size: 15px; color: #333; }
ul, ol { padding-left: 30px; color: #333; }
li { font-size: 15px; color: #333; }
</style>

<p>第一段内容...</p>
<br>
<p>第二段内容...</p>
<br>
<h2>一、章节标题</h2>
<br>
<p>段落内容...</p>
<br>
<p>段落内容...</p>
<br>
<h3>小标题</h3>
<br>
<p>段落内容...</p>
```

---

## 六、常见问题

### Q1：为什么我的标题颜色没生效？

可能原因：
1. CSS 里 `color` 属性出现了两次，后面的覆盖前面的
2. `<style>` 标签内的选择器写错了（如 `h2` 写成了 `.h2`）

### Q2：为什么段落之间没有间距？

`_clean_wechat_margins()` 删掉了所有 `margin`。解决方法：用 `<br>` 代替。

### Q3：为什么 `<br>` 插在列表/引用框内部了？

插入 `<br>` 的脚本应该保护嵌套容器。如果发现列表内部多了空行，说明保护逻辑没生效。

### Q4：能不能直接在草稿箱里调？

可以。用 `<br>` 生成的空行，在微信草稿箱 WYSIWYG 编辑器里**能直接编辑**——按回车增加、删掉 `<br>` 减少。

---

## 七、和 `design.md` 的关系

`design.md` 定义了整体排版风格（配色、字体、响应式）。本文档定义的是**微信 API 特有的渲染限制和对应解决方案**。

建议配合使用：

1. 先用 `design.md` 定风格
2. 参照本文档保证在微信里渲染正确
3. 推送时 wechat-oa 自动处理 CSS 内联和 margin 清理
