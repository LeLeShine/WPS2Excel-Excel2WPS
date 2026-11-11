# DESIGN.md — WPS Excel 图片修复工具 界面设计规范

> **设计系统来源**：[VoltAgent/awesome-design-md](https://github.com/VoltAgent/awesome-design-md) →
> [`design-md/apple/DESIGN.md`](https://github.com/VoltAgent/awesome-design-md/blob/main/design-md/apple/DESIGN.md)（Apple · 浅色 · 高级留白）
>
> 本文件是本项目图形界面的**唯一样式基准**：颜色、字体、圆角、间距、组件规则先在此定义，
> 再映射为 `wps_repair_standalone.py` 中的 `C / FONT / RADIUS / SPACE` 常量。
> 修改界面外观时，先改本文件，再改代码。

---

## 1. 色彩令牌（colors）

唯一交互强调色 = **Action Blue `#0066cc`**，所有"可点击"元素只能用它。

| 令牌 | 值 | 用途（本项目） |
|---|---|---|
| `colors.primary` | `#0066cc` | 主按钮、文本链接、复选框选中、进度条填充、列表选中行 |
| `colors.primary-focus` | `#0071e3` | 按钮悬停/按下、链接悬停、拖放区激活描边 |
| `colors.ink` / `body` | `#1d1d1f` | 标题、正文、日志文字、复选框标签 |
| `colors.ink-muted-80` | `#333333` | 次要正文 |
| `colors.ink-muted-48` | `#7a7a7a` | 副标题、状态提示、拖放区说明、清空类弱提示 |
| `colors.hairline` | `#e0e0e0` | 卡片 1px 描边 |
| `colors.divider-soft` | `#f0f0f0` | 进度条轨道 |
| `colors.canvas` | `#ffffff` | 窗口底色、卡片表面 |
| `colors.canvas-parchment` | `#f5f5f7` | 拖放区表面（以表面色而非描边做区分） |
| `colors.surface-pearl` | `#fafafc` | 珍珠白备用表面 |
| `colors.surface-chip-translucent` | `#d2d2d7` | 禁用态按钮底、未选中复选框描边 |
| `colors.on-primary` | `#ffffff` | 蓝底上的文字/对勾 |
| ~~第二强调色~~ | — | **禁止**（唯一例外见 §5 语义扩展） |

## 2. 字体令牌（typography）

原规范为 Web（SF Pro + CSS px）。桌面 Tk 映射规则：

- **尺寸**：CSS px × 0.75 = Tk 磅值（pt）（96 DPI 基准）
- **字重梯**：300 / 400 / 600 / 700，**没有 500**；正文恒为 400，强调为 600
- **字体替换**（原规范"非 Apple 平台"条款）：SF Pro → **Microsoft YaHei UI**（Windows 系统 UI 字体，含字重 Light/Regular/Bold）；日志等宽 SF Mono → **Consolas**
- Tk 无法表达 600 字重：600 档映射为 `bold`（已知偏差，见 §5）
- Tk 无法表达字距（letter-spacing）：标题的负字距（Apple tight）不实现（已知偏差，见 §5）

| 令牌 | 原规范 | 本项目映射 | 用途 |
|---|---|---|---|
| `typography.tagline` | 21px / 600 | YaHei UI **16pt bold** | 窗口主标题 |
| `typography.body` | 17px / 400 | YaHei UI **13pt** | 按钮文字、拖放区主文案基准 |
| `typography.body-strong` | 17px / 600 | YaHei UI **13pt bold** | 拖放区主文案、**板块标题**（WPS → Excel 修复 / Excel → WPS 修复） |
| `typography.caption` | 14px / 400 | YaHei UI **11pt** | 副标题、列表行、状态、链接、复选框 |
| `typography.caption-strong` | 14px / 600 | YaHei UI **11pt bold** | 卡片小标题、计数、百分比 |
| `typography.button` | 引用 body | YaHei UI **13pt** | 主按钮文字 |
| （等宽扩展） | SF Mono → | **Consolas 10pt** | 处理日志 |

## 3. 圆角语法（rounded）与间距（spacing）

**圆角不混用**：紧凑工具 = `sm`，卡片 = `lg`，"行动" = `pill`，此外没有中间值。

| 圆角令牌 | 值 | 本项目 |
|---|---|---|
| `rounded.xs` | 5px | 复选框方块 |
| `rounded.sm` | 8px | （预留：紧凑工具件） |
| `rounded.md` | 11px | （预留：Pearl 按钮） |
| `rounded.lg` | 18px | 文件卡片、日志卡片、拖放区 |
| `rounded.pill` | 9999px | 主按钮、进度条 |

| 间距令牌 | 值 | 本项目 |
|---|---|---|
| `spacing.xxs` / `xs` / `sm` | 4 / 8 / 12px | 行内小间距、滚动条边距、卡片头部与内容间距、板块标题与拖放区间距 |
| `spacing.md` / `lg` / `xl` | 17 / 24 / 32px | 区块间距 / 卡片内边距 24 / 窗口外边距 32、**左右双板块间距 32** |

### 整体布局（双板块）

页眉（tagline 主标题 + caption 副标题）之下，**左右并列两个等宽修复板块**（窗口 1160×760，
每板块 ~532px），板块间仅用 `spacing.xl` 32px 留白分隔（Apple 规则：分隔靠留白/表面色，不加线不换色）。
每个板块自上而下结构一致：板块标题（body-strong + caption 说明）→ 拖放区（parchment）
→ 文件卡片（可伸展）→ 控制行（胶囊主按钮 + 复选框）→ 进度（percent + pill 条 + 状态）
→ 日志卡片（可伸展）。两块各自独立运行，互不共享状态。

## 4. 组件映射（components）

| 本项目控件 | Apple 组件依据 | 实现要点 |
|---|---|---|
| `RepairPanel` 板块 | `store-utility-card` 组合 + 留白分隔 | 双板块并列容器；板块标题 body-strong，说明 caption-muted；内部结构见 §3 整体布局 |
| `RoundedCard` | `store-utility-card` | 白底 + **1px hairline** + `rounded.lg` 18 + 内边距 24；**无阴影** |
| 拖放区 | `product-tile-parchment` + `configurator-option-chip-selected` | `canvas-parchment` 表面色（色彩变化即分隔，不加描边）；拖入悬停时描边升级为 **2px primary-focus** |
| `AppleButton` | `button-primary` | `primary` 底 + 白字 + **pill 全胶囊** + 内边距 11×22；按下触发系统微交互 **scale(0.95)**；禁用态 `chip` 底 + `ink-muted-48` 字 |
| `LinkLabel` | `text-link` | 透明底 + `primary` 文字，无边框（列表头部操作、清空日志） |
| `AppleCheckbox` | 复选语法（`rounded.xs` + primary） | 未选中：白底 1px `chip` 描边（悬停升 `primary`）；选中：`primary` 填充 + 白色对勾 |
| `AppleProgress` | pill 语法扩展（§5） | 轨道 `divider-soft`，填充 `primary`，胶囊圆角 |
| 列表选中行 | 单一强调色规则 | `primary` 底 + `on-primary` 白字 |
| 处理日志 | `footer` 等宽扩展（§5） | 白底、`ink` 文字、错误行 `error`（§5） |

## 5. 本项目的已知偏差与扩展（相对原规范）

1. **语义错误色 `#d70015`**：原规范分析的是 apple.com 营销页，未涵盖表单/错误状态（"Known Gaps" 已注明）。
   本项目仅在日志**错误反馈**中使用红色；它不是"可点击"信号，不构成第二强调色。
2. **等宽日志字体**：规范无 mono 令牌，按"非 Apple 平台替换"条款 SF Mono → Consolas。
3. **进度条**：规范无该组件，按 pill 语法与唯一强调色规则实现（轨道 `divider-soft` + 填充 `primary`）。
4. **字重 600 → bold**：Tk 字体模型无 600，`bold`(700) 近似；浅色 UI 下视觉等效。
5. **字距不实现**：Tk 无 letter-spacing；标题"Apple tight"负字距放弃。
6. **动效**：按下 `scale(0.95)` 以重绘实现；hover 仅用 `primary-focus`（规范不记录 hover，桌面端保留最小悬停反馈）。

## 6. Do / Don't（执行清单）

- ✅ 所有"可点击"元素只用 `primary #0066cc`；✅ 标题/卡片用 `lg 18`，行动用 `pill`，复选框用 `xs 5`，不混用
- ✅ 卡片只有 1px hairline，**无阴影、无渐变**（阴影规范仅留给产品图片，本项目无产品图）
- ✅ 正文 400、强调 600（bold），不用 500；✅ 留白优先：外边距 32、区块 17、卡片内 24
- ❌ 不引入第二强调色；❌ 不给卡片/按钮/文字加阴影；❌ 不用装饰性渐变背景

## 7. 迭代方式

1. 在原规范（awesome-design-md → apple）中选定令牌 → 更新本文件对应表格
2. 同步 `wps_repair_standalone.py` 顶部常量区（`C` / `FONT` / `RADIUS` / `SPACE`）
3. 截图回归：拖放区、卡片、按钮胶囊、进度条、日志五个区块逐一对照本文件；
   双板块另查：左右等宽（各 ~532px）、各区块横向对齐、板块间距 32、无第三色出现
