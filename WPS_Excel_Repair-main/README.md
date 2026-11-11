# WPS Excel 图片修复工具

**双向原生转换：WPS 的 `=_xlfn.DISPIMG()` 单元格图片 → Excel 原生「置于单元格」图片；
Excel 的「图片置于单元格」图片 → WPS 原生单元格图片。图片字节原样搬运，不重新编码。**

## 📋 项目背景

### 问题一：WPS 写的文件在 Excel 中不显示图片
WPS Office 创建的 Excel 文件用 `=_xlfn.DISPIMG("ID_xxx",1)` 公式 + `xl/cellimages.xml`
存储「嵌入单元格图片」。这种格式在 Microsoft Excel 中：

- 图片完全不显示（单元格只剩公式文本）
- 原始图片无法通过常规方式取出

### 问题二：Excel 写的文件在 WPS 中不显示图片
Excel 2024 / Microsoft 365 的「**图片置于单元格**」（Place in Cell）把图片存为 rich value
（`xl/metadata.xml` + `xl/richData/*` + 单元格 `vm` 属性），这种格式在 WPS 中：

- 图片完全不显示（单元格只剩 `#VALUE!` 错误）
- 原始图片无法通过常规方式取出

### 为什么是「原生转原生」
两家的「图片在单元格里」都是**厂商私有结构**，不是普通浮动图片。浮动图片锚点
（`twoCellAnchor` 等）**无法表达**它们的动态行为：

| | WPS 嵌入单元格图片 | Excel 置于单元格图片 |
|---|---|---|
| 随格移动 | ✅ | ✅ |
| 缩小 | 随格等比缩小 | 随格等比缩小 |
| 放大 | **以置入时大小为上限**，不无限放大 | **动态取 min(格宽/图宽, 格高/图高)**，随格变大 |

因此本工具把每张图片改写成**目标软件自己的原生置入结构**——转换结果在目标软件里
与该软件自己插入的置入图片行为完全一致（已实机逐像素验证，见下文）。

## 🔧 技术原理

### 方向一：WPS DISPIMG → Excel 置于单元格（rich value 全链路）
1. **解析**：`xl/cellimages.xml` + `xl/_rels/cellimages.xml.rels` 把每个 DISPIMG
   公式单元格还原为具体图片部件（同图多格自动去重，共享一个媒体文件）
2. **改写**：删除 DISPIMG 公式；写入 Excel rich value 全链路——
   - `xl/metadata.xml`（futureMetadata/valueMetadata 槽位）
   - `xl/richData/rdrichvalue.xml` + `rdrichvaluestructure.xml` + `rdRichValueTypes.xml`
   - `xl/richData/richValueRel.xml` + `_rels`（指向原样搬运的图片字节）
   - 5 条 workbook 关系 + 4 个 Content_Types Override
   - 单元格 `<c t="e" vm="N"><v>#VALUE!</v></c>`（`vm` 为 1 基序号）
3. **行为**：Excel 原生等比自适应——`scale = min(格宽/图宽, 格高/图高)`，行高列宽
   变化时动态重算；非填满时贴单元格左下角（Excel 默认对齐）。

> 实机验证（Excel 2024）：6 张图全部按 min 规则渲染、左/下贴边、比例不变；
> 行高列宽 ×2 时精确 ×2 缩放；文件经 Excel 重新保存后 rich value 链路完整保留
> （Excel 原生接受，未做静默修复）；`Shapes.Count = 0`（确为置入而非浮动对象）。

### 方向二：Excel 置于单元格 → WPS DISPIMG/cellimages
1. **解析**：按 `vm` → metadata 槽位 → `rdrichvalue` → `richValueRel` → 媒体的链路
   还原每格图片；清空 rich value 引用（`vm` 属性与 metadata/richData 部件随重建移除）
2. **改写**：写入 WPS 原生结构（与 WPS 自己保存的文件逐字段一致）——
   - `xl/cellimages.xml`：`etc:cellImage` 完整条目（`cNvPr` name=`ID_<32位大写十六进制>`、
     `cNvPicPr/picLocks`、`blipFill/stretch/fillRect`、`spPr/xfrm + prstGeom rect`）
   - `xl/_rels/cellimages.xml.rels` + 原样搬运的图片字节
   - Content_Types：`application/vnd.wps-officedocument.cellimage+xml`
   - **workbook.xml.rels 关系** `Type="http://www.wps.cn/officeDocument/2020/cellImage"`
     ——缺此关系 WPS 会静默忽略整个部件（实测图片全都不渲染）
   - 单元格公式 `=_xlfn.DISPIMG("ID_xxx",1)`，`t="str"` + 缓存值 `=DISPIMG("ID_xxx",1)`
3. **行为**：WPS 原生嵌入单元格行为——随格移动、等比 contain 适配（贴合限轴、
   松边居中）、缩小随格等比、**放大约束到置入大小**（实测上限 = 图片原始尺寸 ×
   屏幕 DPI 缩放）。

> 实机验证（WPS）：3 张图全部渲染、contain 规则逐格吻合；行高列宽 ×0.5 时精确
> ×0.5 等比缩小；×2 时放大封顶于原生尺寸×DPI（125% 缩放下实测 scale=1.25，
> 而非 min 规则的 1.6）；`Shapes.Count = 0`（确为置入而非浮动对象）；WPS 重新
> 保存正常。
>
> 另注：WPS 会**自动把含置入图片的未自定义行高行撑高到图片所需高度**
> （实测默认行 13.8pt 被撑到 48.6pt），与 WPS 自己插入置入图片时表现一致；
> 输出转换前的单元格行高列宽不变。

> **格式来源**：方向二输出以真实 WPS 文件（用户文档
> 《土山汉墓遗址展陈文物选配清单.xlsx》）为基准逐字段对齐。此前的简化结构
> （ContentType 误写为 application/xml、缺 workbook 关系、公式单参数且无缓存值）
> 会导致 WPS 完全不渲染——这四点已全部修正并有回归测试保护。

### 原样字节（两个方向一致）
图片媒体**逐字节原样搬运**（zip 内直接读写），不经过任何解码/重编码，
PNG/JPG 的压缩数据与原文件完全相同；扩展名保留。

### ⚠️ 跨厂商注意（固有限制，非缺陷）
转换结果是**目标厂商的私有置入格式**，源厂商打不开它：

- 方向一输出（Excel 置于单元格）**在 WPS 中不显示**——WPS 不支持 rich value
- 方向二输出（WPS DISPIMG）**在 Excel 中不显示**——Excel 不支持 cellimages

需要在另一家查看时，把输出文件拖回界面对应的另一块面板转回去即可（往返无损，
媒体字节不变）。

## 🚀 功能特性

### ✅ 已完成特性
- **双向原生转换**：界面左右两块并列——左「WPS → Excel」，右「Excel → WPS」
- **方向自动判别**：按包内容识别（`cellimages.xml` → 左；`metadata.xml`+`richData` → 右）
- **拖拽操作**：支持文件拖拽到程序图标，或拖放到界面左右对应板块
- **进度显示**：每个板块独立的实时进度与处理日志，可同时运行互不干扰
- **图片字节原样保留**：不重新编码，压缩数据与原文件逐字节相同
- **同图去重**：同一图片被多个单元格引用时只搬运一份媒体
- **批量处理**：支持多个工作表、多个文件同时处理
- **自动保存**：生成新文件，保留原始文件
- **自动打开**：修复完成后自动打开文件（每块可独立勾选）

### 🎯 精确转换
- **完整链路**：两个方向都按对方厂商的真实文件结构逐字段写入（含 Content_Types、
  workbook 关系、缓存公式值等易漏项）
- **边界保护**：`vm` 序号、关系 ID、部件命名均按 OOXML 规则分配，不冲突
- **兜底方案**：个别单元格失败时记录 `[FAIL]` 继续处理其余单元格

## 🎨 界面设计规范

图形界面严格遵循 **[awesome-design-md](https://github.com/VoltAgent/awesome-design-md)** 收录的
**Apple 设计系统**（`design-md/apple/DESIGN.md`），令牌与映射见项目根目录 **[DESIGN.md](DESIGN.md)**：

- **单一强调色**：全部可点击元素只用 Action Blue `#0066cc`（悬停 `#0071e3`）
- **浅色高级留白**：纯白画布 `#ffffff` + 羊皮纸表面 `#f5f5f7`，窗口外边距 32px、区块间距 17px、卡片内边距 24px
- **字体层级**：正文 17px→13pt / 说明 14px→11pt（CSS px × 0.75 = Tk 磅值），字重梯 400/600，无 500
- **圆角语法**：卡片 `18px`、复选框 `5px`、主按钮与进度条全胶囊 `pill`，不混用
- **无阴影无渐变**：卡片只有 1px `#e0e0e0` 细线描边，层次靠表面色区分
- **组件实现**：胶囊主按钮按下带 `scale(0.95)` 微交互；拖放区悬停升级为 2px 蓝色描边

## 🛠️ 使用方法

### 方式一：图形界面（推荐）
1. 双击运行 `WPS_Excel_Repair_Tool.exe` 打开界面（左右两个板块）
2. **添加文件**：把 .xlsx 文件拖放到**对应板块**（WPS 文件放左边、Excel 置于单元格图片文件放右边，
   拖到窗口其他位置会按文件内容自动分流），或点击板块内区域选择（支持多选、批量）
3. **开始处理**：点击该板块的「开始修复」按钮（两块可先后运行，互不影响）
4. **原位置保存**：结果文件保存在**源文件所在文件夹**，命名为 `原文件名_fixed.xlsx`（不覆盖原文件）
5. 勾选「修复完成后自动打开文件」可在处理完后自动打开结果（每块独立勾选）

界面还提供：文件列表管理（移除选中/清空）、独立进度条与处理日志、失败原因提示。

### 方式二：拖拽到程序图标
1. 将一个或多个 .xlsx 文件拖拽到 `WPS_Excel_Repair_Tool.exe` 图标上
2. 文件按内容自动判别方向并进入对应板块，再点击「开始修复」即可

### 方式三：无界面批处理（命令行）
```bash
WPS_Excel_Repair_Tool.exe --batch 文件1.xlsx 文件2.xlsx
```
每个文件按包内容自动判别方向，日志以 `[WPS->Excel]` / `[Excel->WPS]` 标注本次采用的方向。

### 方式四：自动开始（无人值守）
```bash
WPS_Excel_Repair_Tool.exe --autostart 文件1.xlsx 文件2.xlsx
```
打开界面并自动对已加载文件的板块点击「开始修复」，适合计划任务等自动化场景。

### 方式五：代码集成
```python
from wps_repair_standalone import PreciseSafeWPSExcelFixer, ReverseCellImageFixer

# 方向一：WPS DISPIMG -> Excel 原生置于单元格图片
fixer = PreciseSafeWPSExcelFixer('wps_file.xlsx')
fixed_file = fixer.fix_excel_file_precise_safe('wps_file_fixed.xlsx')

# 方向二：Excel 置于单元格图片 -> WPS 原生单元格图片
fixer2 = ReverseCellImageFixer('excel_file.xlsx')
converted = fixer2.convert_excel_to_wps('excel_file_fixed.xlsx')
```

### 从源码运行
```bash
pip install -r requirements.txt
python wps_repair_standalone.py            # 打开图形界面
python wps_repair_standalone.py --batch a.xlsx b.xlsx   # 批处理
```

### 打包EXE
```bash
python build_exe.py
# 生成 WPS_Excel_Repair_Tool.exe（单文件、无控制台窗口）
```

## 📊 性能表现

| 文件大小 | 工作表数量 | 图片数量 | 修复时间 |
|---------|------------|----------|----------|
| 8.7MB   | 16         | 100+     | ~15秒    |
| 2MB     | 5          | 20       | ~3秒     |
| 500KB   | 2          | 5        | ~1秒     |

**优化措施：**
- 延迟加载：按需加载工作表
- 进度显示：避免用户误以为程序卡死
- 内存管理：及时释放大文件内存

## 🔍 调试与故障排除

### 常见问题
1. **文件损坏警告**：确保使用最新版本；两个方向的输出都经对应软件实机打开 + 重新保存验证
2. **方向一输出在 WPS 里看不到图片**：固有限制——WPS 不支持 Excel rich value，
   需要给 WPS 用时请走方向二（把原 WPS 文件转成 Excel 置于单元格，或把 Excel 文件转回 WPS 格式）
3. **方向二输出在 Excel 里看不到图片**：固有限制——Excel 不支持 WPS cellimages，
   需要给 Excel 用时请走方向一
4. **Excel 单元格显示 `#VALUE!`**：这是 rich value 占位的正常形态，图片在单元格内正常显示
   （Excel 自己创建的置于单元格图片也是同样表现）

### 日志输出
程序会在修复过程中输出详细日志：
```
=== WPS 单元格图片 -> Excel「置于单元格」图片 ===
发现 6 个DISPIMG公式
收集完成: 6 个单元格, 6 个独立图片（原样保留）
  媒体: IMGID_1 -> xl/media/wpscellimg1.png (1946 字节, 不重新编码)
```

**注意**：本工具处理两个方向的兼容问题——WPS 创建的 DISPIMG 文件与
Excel 创建的「置于单元格图片」文件；普通 Excel 浮动图片文件两边都能正常打开，
无需处理。
