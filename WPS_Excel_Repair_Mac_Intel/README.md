# WPS Excel 图片修复工具 · Intel Mac 版

**双向原生转换：WPS `=_xlfn.DISPIMG()` 单元格图片 ↔ Excel「置于单元格」rich value 图片。
图片字节原样搬运，不重新编码。** 本文件夹是 macOS (Intel / x86_64) 适配版。

> 转换原理、双向行为规范与实机验证细节与 Windows 版完全一致，见
> `../WPS_Excel_Repair-main/README.md`（主 README）与本目录 `DESIGN.md`。
> 转换核心代码逐字节同源（纯 `zipfile`/`XML`/`openpyxl`，无任何平台调用）。

## 与 Windows 版的差异（仅 4 处平台适配）

| # | 点 | Windows | 本 Mac 版 |
|---|---|---|---|
| 1 | 字体族 | Microsoft YaHei UI / Consolas | **PingFang SC / Menlo**（系统自带，CJK 正常渲染） |
| 2 | 自动打开结果 | `os.startfile` | macOS **`open`**（Finder 打开） |
| 3 | 文件入口 | 拖到程序图标 → argv | 拖到 **Dock 图标** → Tk 原生 `::tk::mac::OpenDocument` 回调（零第三方依赖） |
| 4 | 打包 | `build_exe.py` → `.exe` | **`build_exe_mac.py` → `.app`**（强制 `--target-architecture=x86_64`） |

窗口内板块拖放（tkinterdnd2）两平台同源；若该组件在你的环境不可用，
会自动降级为「点击板块区域选择文件」（界面提示会同步变化），核心功能不受影响。

> 已知环境限制：Python 3.13+ 自带 Tcl/Tk 9.0，而 tkinterdnd2 在 Intel Mac 上
> 只发布了 Tcl 8.x 的 tkdnd 二进制，窗口内拖放会自动降级（拖到 **Dock 图标**
> 与点击选择仍可用）。要窗口内拖放请用 Python 3.10~3.12。

## 📦 在 Intel Mac 上打包

前置：macOS + Python 3.10~3.14（`python3 --version`），带 Tcl/Tk 的官方安装包即可。

```bash
cd WPS_Excel_Repair_Mac_Intel
python3 -m pip install -r requirements.txt   # 脚本也会自动补装缺失依赖
python3 build_exe_mac.py
```

打包脚本自带三道闸，**不会再产出“打开即闪退”的空壳 App**：

1. **依赖预检**：确认 `openpyxl / tkinter / Pillow / tkinterdnd2` 可导入，
   缺失自动 `pip install -r requirements.txt`，仍缺失则终止并给出提示；
2. **告警扫描**：解析 PyInstaller 的 `warn-*.txt`，入口脚本关键模块缺失直接判失败；
3. **启动自检**：对产物分别执行 `--selfcheck`（界面初始化）与 `--batch`（批处理入口），
   启动即崩的产物在这里就会被判失败。

产物：`WPS_Excel_Repair_Tool.app`（**目录模式 onedir**，约 **30 MB**，x86_64；
独立运行，用户无需安装任何依赖）。

- 采用目录模式而非单文件模式：启动为**单进程直接加载**，Dock 图标只出现一次、
  窗口很快出现。单文件模式（onefile）会先跑一个解压用的父进程，表现为图标
  「出现→消失→再出现」、界面迟迟不出现——故本项目已固定用 onedir。
- **体积优化**：打包时排除了 openpyxl 的三组「可选加速 / 可选功能」依赖
  （`numpy` 约 41 MB、`Pillow` 约 9 MB、`lxml` 约 8 MB），运行时 openpyxl
  自动回退到标准库 ElementTree、不使用图片像素接口——功能不变，
  这些模块既不会被打进 `.app`，也不需要装在系统里。
- **分发**：发给他人时请右键 `.app` → 压缩，对方解压后同样右键 → 打开。
- 脚本强制 **x86_64（Intel）**；若用 Apple Silicon 的原生 Python 无法交叉编译，会自动回退并打印警告——请改在 Intel Mac 上重跑。
- **首次打开（Gatekeeper 隔离属性）**：从网络拷来的 `.app` 需右键 → 打开，或执行：
  ```bash
  xattr -dr com.apple.quarantine WPS_Excel_Repair_Tool.app
  ```

## 🛠️ 使用方法

### 图形界面（推荐）
1. 双击 `WPS_Excel_Repair_Tool.app` 打开（左右两个板块）。
2. **添加文件**（三种方式任选）：
   - 把 `.xlsx` 拖到**窗口内对应板块**（WPS 文件放左、Excel 置于单元格图片文件放右，拖到板块外会按内容自动分流）；
   - 把文件拖到 **Dock 程序图标**（自动按包内容分流到对应板块）；
   - 点击板块内区域选择（支持多选、批量）。
3. 点击该板块「开始修复」（两块可先后运行，互不影响）。
4. 结果保存在**源文件所在文件夹**，命名 `原文件名_fixed.xlsx`（不覆盖原文件）；勾选「修复完成后自动打开文件」会用 Finder 打开结果。

### 命令行批处理
```bash
# .app 包内可执行体（终端里能看到日志与退出码）
WPS_Excel_Repair_Tool.app/Contents/MacOS/WPS_Excel_Repair_Tool --batch 文件1.xlsx 文件2.xlsx

# 或直接跑源码
python3 wps_repair_standalone.py --batch 文件1.xlsx 文件2.xlsx

# 无人值守：打开界面并自动开始
python3 wps_repair_standalone.py --autostart 文件1.xlsx

# 启动自检（只构建界面不进主循环，打包脚本也用它做冒烟测试）
python3 wps_repair_standalone.py --selfcheck
```

### 代码集成
```python
from wps_repair_standalone import PreciseSafeWPSExcelFixer, ReverseCellImageFixer

# 方向一：WPS DISPIMG -> Excel 原生置于单元格图片
PreciseSafeWPSExcelFixer('wps_file.xlsx').fix_excel_file_precise_safe('wps_file_fixed.xlsx')

# 方向二：Excel 置于单元格图片 -> WPS 原生单元格图片
ReverseCellImageFixer('excel_file.xlsx').convert_excel_to_wps('excel_file_fixed.xlsx')
```

## ⚠️ 跨厂商注意（固有限制，两个平台相同）

转换结果是**目标厂商的私有置入格式**，源厂商打不开它：

- 方向一输出（Excel 置于单元格 rich value）**在 WPS 中不显示**；
- 方向二输出（WPS DISPIMG/cellimages）**在 Excel 中不显示**。

需要在另一家查看时，把输出文件拖回界面对应的另一块面板转回去即可（往返无损，媒体字节不变）。

> 本工具只影响“单元格内图片”的存储格式；普通浮动图片文件两个软件都能直接打开，无需处理。
> Excel 单元格显示 `#VALUE!` 是 rich value 占位的正常形态（Excel 自建的置于单元格图片同样如此），图片在单元格内正常显示。

## 🧯 故障排除

| 现象 | 原因 / 处理 |
|---|---|
| Dock 图标「出现→消失→再出现」 | 旧版**单文件（onefile）**打包的固有行为：先由父进程把 37MB 解压到临时目录，父进程退出后真程序才接管。现打包脚本已固定用**目录模式（onedir）**，重新执行 `python3 build_exe_mac.py` 即可，图标只出现一次 |
| 双击 `.app` 打开即退 | 程序已内置启动守护：会弹出错误框说明原因，并把完整堆栈写入日志（不再静默退出）。最常见原因是**打包时缺依赖**（缺 `openpyxl` 时打包器只告警不报错，产物启动即 `ModuleNotFoundError`）：在本目录执行 `python3 -m pip install -r requirements.txt` 后重新 `python3 build_exe_mac.py`。新版打包脚本会自动预检依赖并做启动自检，不会再产出这种产物 |
| 提示“无法打开 / 已损坏” | Gatekeeper 隔离属性：`xattr -dr com.apple.quarantine WPS_Excel_Repair_Tool.app` |
| 窗口内拖不进文件 | Python 3.13+（Tcl/Tk 9）在 Intel Mac 上没有对应的 tkdnd 构建，拖放自动降级；改用拖到 **Dock 图标** 或点击板块内区域选择 |
| 需要排查问题 | 日志：`~/Library/Logs/WPS_Excel_Repair_Tool.log`（Windows：`%LOCALAPPDATA%\WPS_Excel_Repair_Tool\`）；在终端直接运行可看到堆栈与退出码：`WPS_Excel_Repair_Tool.app/Contents/MacOS/WPS_Excel_Repair_Tool` |

## 🎨 界面规范

沿用主项目 Apple 设计系统实现（`DESIGN.md`）：唯一强调色 Action Blue `#0066cc`、
浅色画布 `#ffffff` + 羊皮纸 `#f5f5f7`、17px 细线卡片、胶囊主按钮、无阴影无渐变。
macOS 下字号映射不变（CSS px × 0.75 = Tk pt），仅字体族替换为 PingFang SC / Menlo。

**紧凑版适配**：版式/配色/圆角语法与规范完全一致，仅把字号与间距各降一档
（正文 17→16px、说明 14→13px、页眉 21→17px；间距令牌 ×0.75），
默认窗口 **1080×600**（原 1160×760，最小可缩至 900×500），
拖放区 80→62px、按钮内边距同步收紧——更密、更矮，风格不变。

## 📁 文件清单

| 文件 | 说明 |
|---|---|
| `wps_repair_standalone.py` | 主程序（转换核心同源 + 4 处 macOS 适配 + 启动守护/崩溃日志） |
| `build_exe_mac.py` | Intel (x86_64) 打包脚本（onedir + 依赖预检 + 告警扫描 + 产物启动自检） |
| `requirements.txt` | 依赖（openpyxl / Pillow / tkinterdnd2 / pyinstaller） |
| `DESIGN.md` | Apple 设计系统令牌与映射（与主项目一致） |
| `README.md` | 本说明 |
