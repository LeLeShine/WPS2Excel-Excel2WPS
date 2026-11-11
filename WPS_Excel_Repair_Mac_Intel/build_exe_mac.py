#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS Excel修复工具 - Intel Mac (x86_64) 打包脚本

在 macOS 上执行：
    python3 build_exe_mac.py

要点：
  * 用 --onedir 而非 --onefile：onefile 启动时先跑一个“父进程”解压 37MB 到
    临时目录，父进程退出后真程序才接管——表现为 Dock 图标「出现→消失→
    再出现」、窗口迟迟不出现。onedir 直接从 .app 内加载，单进程启动，
    图标只出现一次，冷启动也快得多。
  * 依赖预检：打包前确认 openpyxl / tkinter / Pillow / tkinterdnd2 可导入，
    缺失时自动执行 pip install -r requirements.txt；仍缺失则终止打包。
    —— 缺依赖正是“.app 双击即闪退”的根因（打包器只告警不报错，产物启动
       就 ModuleNotFoundError，--windowed 下没有任何可见提示）。
  * 强制目标架构 --target-architecture=x86_64（Intel 版）；若当前 Python
    无法交叉编译（如 Apple Silicon 原生 Python），自动回退并打印警告。
  * 解析 PyInstaller 的 warn-*.txt：入口脚本的关键模块缺失直接判失败。
  * 产物冒烟测试：跑 --selfcheck（界面初始化）与 --batch（批处理入口），
    “启动即崩”的产物在这里就会被判失败，不会交付给用户。
  * 产物旁自动清理 build/ 与 .spec 中间文件，并移除与 .app 内容重复的裸产物
    （onedir 时 PyInstaller 会额外产出一份同名目录，属全量副本）。
"""

import glob
import os
import re
import shutil
import subprocess
import sys

# 始终以脚本所在目录为工作目录（允许从任意位置调用）
os.chdir(os.path.dirname(os.path.abspath(__file__)) or '.')

APP = "WPS_Excel_Repair_Tool"
MAIN = "wps_repair_standalone.py"
REQUIREMENTS = "requirements.txt"

# 打包环境必须能导入的模块：缺失 -> 产物启动即崩
CRITICAL_MODULES = ("openpyxl", "tkinter")
# 缺失只降级/影响体验，不阻断打包
OPTIONAL_MODULES = (("PIL", "Pillow"), ("tkinterdnd2", "tkinterdnd2"))
# PyInstaller warn 文件里出现即判失败的顶层模块
WARN_FAIL_MODULES = ("openpyxl", "tkinter", "_tkinter")


def _size_mb(path):
    """真实磁盘占用（PyInstaller onedir 内有大量硬链接，按 inode 去重）"""
    if os.path.isfile(path):
        return os.path.getsize(path) / (1024 * 1024)
    total = 0
    seen = set()
    for root, _dirs, files in os.walk(path):
        for f in files:
            try:
                st = os.stat(os.path.join(root, f))
            except OSError:
                continue
            key = (st.st_dev, st.st_ino)
            if key in seen:
                continue
            seen.add(key)
            total += st.st_size
    return total / (1024 * 1024)


def _run(cmd, timeout=600):
    """执行子进程，统一按 UTF-8 读取输出（系统为 GBK 代码页时也不会崩）"""
    return subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=timeout)


def _importable(module_name):
    """只探测能否导入，不产生窗口/副作用"""
    try:
        __import__(module_name)
        return True
    except Exception:
        return False


def _print_arch(binary):
    """报告产物架构（尽力而为：lipo 不存在时跳过）"""
    try:
        r = subprocess.run(["lipo", "-archs", binary], capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        if r.returncode == 0 and r.stdout.strip():
            print("架构: %s" % r.stdout.strip())
    except Exception:
        pass


def check_dependencies():
    """依赖预检：关键缺失 -> 自动安装 -> 仍缺失则终止（绝不产出启动即崩的产物）"""
    missing_critical = [m for m in CRITICAL_MODULES if not _importable(m)]
    missing_optional = [pip for mod, pip in OPTIONAL_MODULES
                        if not _importable(mod)]

    if (missing_critical or missing_optional) and os.path.exists(REQUIREMENTS):
        print("缺少依赖 %s，正在执行: pip install -r %s"
              % (", ".join(missing_critical + missing_optional), REQUIREMENTS))
        try:
            _run([sys.executable, "-m", "pip", "install", "-r", REQUIREMENTS],
                 timeout=900)
        except Exception as exc:
            print("依赖安装执行失败：%s" % exc)
        missing_critical = [m for m in CRITICAL_MODULES
                            if not _importable(m)]
        missing_optional = [pip for mod, pip in OPTIONAL_MODULES
                            if not _importable(mod)]

    if missing_critical:
        if "tkinter" in missing_critical:
            print("错误：当前 Python 缺少 Tkinter（tk 模块），无法打包图形界面。")
            print("      请安装带 Tcl/Tk 的 Python 3.10~3.12（python.org 官网")
            print("      安装包自带 Tk），或用 `python3 -m pip install -r %s`"
                  % REQUIREMENTS)
        else:
            print("错误：关键依赖仍缺失：%s" % ", ".join(missing_critical))
            print("      请手动执行: python3 -m pip install -r %s" % REQUIREMENTS)
        return False

    if missing_optional:
        print("警告：可选依赖缺失 %s（拖放/图片处理会自动降级，核心转换不受影响）"
              % ", ".join(missing_optional))
    else:
        print("依赖预检通过：openpyxl / tkinter / Pillow / tkinterdnd2 均可用")
    return True


def scan_pyinstaller_warnings():
    """解析 PyInstaller 的 warn-*.txt：入口脚本的关键模块缺失 -> 判失败"""
    hard, soft = [], []
    for warn_file in glob.glob(os.path.join("build", "*", "warn-*.txt")):
        try:
            with open(warn_file, encoding="utf-8", errors="replace") as handle:
                lines = handle.readlines()
        except OSError:
            continue
        for line in lines:
            matched = re.match(r"\s*missing module named '?([\w.]+)'? - ",
                               line)
            if not matched:
                continue
            name = matched.group(1)
            if name in WARN_FAIL_MODULES:   # 精确匹配，openpyxl.tests 之类不误伤
                hard.append(name)
            elif name.startswith(("tkinterdnd2", "PIL")):
                soft.append(name)
    return sorted(set(hard)), sorted(set(soft))


def _is_display_error(output):
    """无图形会话（SSH/CI）时 Tk 起不来，这类失败不该判打包失败"""
    markers = ("display name and and screen", "can't connect to display",
               "no display name", "couldn't connect to display",
               "invalid display name", "Aqua")
    return any(marker in output for marker in markers)


def smoke_test(binary, timeout=90):
    """启动自检：--selfcheck（界面初始化）+ --batch（批处理入口）。
    返回 (ok, messages)：成功时 messages 为警告；失败时最后一条是失败原因。
    “启动即崩”的产物在这里就会暴露。"""
    messages = []
    checks = ((["--selfcheck"], 0), (["--batch"], 2))
    for args, expected in checks:
        try:
            result = _run([binary] + args, timeout=timeout)
        except subprocess.TimeoutExpired:
            return False, messages + ["%s 运行超时（可能卡在启动）"
                                      % " ".join(args)]
        except OSError as exc:
            return False, messages + ["无法执行产物：%s" % exc]

        output = (result.stdout or "") + (result.stderr or "")
        if result.returncode != expected:
            if args[0] == "--selfcheck" and _is_display_error(output):
                messages.append("当前会话无图形显示，跳过界面自检（--selfcheck）")
                continue
            return False, messages + [
                "%s 期望退出码 %d，实际 %d\n%s"
                % (" ".join(args), expected, result.returncode,
                   output[-1200:])]
        if "Traceback" in output or "ModuleNotFoundError" in output:
            return False, messages + ["%s 输出含异常：\n%s"
                                      % (" ".join(args), output[-1200:])]
        if args[0] == "--selfcheck" and "selfcheck ok" not in output:
            return False, messages + ["--selfcheck 未输出成功标记：\n%s"
                                      % output[-1200:]]
    return True, messages


def drop_redundant_product(app_binary, bare_binary):
    """PyInstaller 会同时产出 .app 包与等价的裸产物：
    onefile 时是同名可执行文件，onedir 时是同名目录（_internal/ 全量副本）。
    .app 已通过自检时移除重复项，避免同一份内容交付两遍。"""
    if not (os.path.isfile(app_binary) and os.path.exists(bare_binary)):
        return 0
    try:
        if os.path.isdir(bare_binary):
            size = int(_size_mb(bare_binary) * 1024 * 1024)
            shutil.rmtree(bare_binary, ignore_errors=True)
            return size if not os.path.exists(bare_binary) else 0
        if os.path.samefile(app_binary, bare_binary):
            return 0
        size = os.path.getsize(bare_binary)
        os.unlink(bare_binary)
        return size
    except OSError:
        return 0  # 删除失败就保留，不影响使用


def build_exe():
    print("=== WPS Excel修复工具 · Intel Mac (x86_64) 打包 ===")
    if sys.platform != "darwin":
        print("错误：此脚本需在 macOS 上运行（当前系统: %s）。" % sys.platform)
        print("请把本文件夹拷贝到 Intel Mac 上执行: python3 build_exe_mac.py")
        return False

    # 1) 依赖预检（缺依赖 = 产物打开闪退的根因）
    if not check_dependencies():
        return False

    # 2) PyInstaller 检查（缺失时自动安装）
    try:
        import PyInstaller
        print("PyInstaller 版本: %s" % PyInstaller.__version__)
    except ImportError:
        print("未检测到 PyInstaller，正在安装...")
        try:
            _run([sys.executable, "-m", "pip", "install", "-U", "pyinstaller"],
                 timeout=900)
        except Exception as exc:
            print("安装 PyInstaller 失败：%s" % exc)
            return False

    # 3) 清理旧构建产物
    for folder in ("build", "dist", "__pycache__"):
        if os.path.exists(folder):
            shutil.rmtree(folder, ignore_errors=True)
            print("删除旧目录: %s" % folder)

    if not os.path.exists(MAIN):
        print("缺少必要文件: %s" % MAIN)
        return False
    print("检测到: %s" % MAIN)

    base_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",           # 目录模式：单进程启动，Dock 图标不会闪两次
        "--windowed",         # 无控制台（macOS 下产出 .app 包）
        "--noconfirm",        # 覆盖不提示
        "--strip",            # 去掉二进制调试符号（不影响运行，显著减小体积）
        # 以下模块都是 openpyxl 的「可选加速/可选功能」依赖，源码里全部
        # try/except 保护（本工具也从不使用对应功能），排除后运行时自动回退，
        # 产物从 ~90MB 降到 ~34MB，且无需用户安装任何额外依赖：
        #   numpy(~41MB) — compat/numbers.py 可选数值类型扩展
        #   lxml(~8MB)   — openpyxl 回退到标准库 ElementTree 解析
        #   PIL(~9MB)    — openpyxl/drawing/image.py 可选；本工具不碰图片像素
        #   ssl(~1MB)    — 运行期无人导入（uuid→hashlib 用的 libcrypto 保留）
        "--exclude-module=numpy",
        "--exclude-module=lxml",
        "--exclude-module=PIL",
        "--exclude-module=ssl",
        "--name=%s" % APP,
        "--distpath=.",
    ]
    # 可选图标（存在才打包）
    if os.path.exists("assset/icon.icns"):
        base_cmd.append("--icon=assset/icon.icns")

    # 4) 第一次尝试强制 Intel (x86_64)；失败则回退本机架构
    attempts = [
        ["--target-architecture=x86_64"],
        [],
    ]
    forced_ok = False
    for extra in attempts:
        cmd = base_cmd + extra + [MAIN]
        print("\n执行: %s" % " ".join(cmd[1:]))
        try:
            r = _run(cmd, timeout=1800)
        except Exception as exc:
            print("打包执行失败：%s" % exc)
            continue
        if r.returncode == 0:
            forced_ok = bool(extra)
            break
        print("尝试失败（末尾输出）:")
        print((r.stderr or r.stdout or "")[-800:])
    else:
        print("打包失败：两种参数组合均未成功")
        return False

    # 5) 解析 PyInstaller 警告：入口脚本关键模块缺失 -> 判失败
    hard_missing, soft_missing = scan_pyinstaller_warnings()
    if soft_missing:
        print("警告：PyInstaller 报告可选模块未找到：%s（功能自动降级）"
              % ", ".join(soft_missing))
    if hard_missing:
        print("打包失败：PyInstaller 报告入口脚本关键模块未打包：%s"
              % ", ".join(hard_missing))
        print("        产物启动会立即崩溃（打开闪退），已终止交付。")
        shutil.rmtree("build", ignore_errors=True)
        return False

    # 6) 识别产物：.app 包 或 裸可执行文件
    app_dir = APP + ".app"
    app_binary = os.path.join(app_dir, "Contents", "MacOS", APP)
    bare = APP
    product = None
    if os.path.isdir(app_dir) and os.path.exists(app_binary):
        product = app_dir
        binary = app_binary
    elif os.path.exists(bare):
        product = bare
        binary = bare
    if product is None:
        print("打包失败：未找到产物（%s / %s）" % (app_dir, bare))
        shutil.rmtree("build", ignore_errors=True)
        return False

    # 7) 冒烟测试：启动即崩的产物在这里就会被判失败
    print("\n冒烟测试: %s" % binary)
    ok, messages = smoke_test(binary)
    if not ok:
        print("打包失败：产物启动自检未通过（即用户会遇到的“打开闪退”）")
        for message in messages:
            print("  %s" % message)
        shutil.rmtree("build", ignore_errors=True)
        return False
    for message in messages:
        print("  警告: %s" % message)
    print("  自检通过：--selfcheck（界面初始化）与 --batch（批处理入口）均正常")

    # 8) 清理中间文件
    shutil.rmtree("build", ignore_errors=True)
    spec = APP + ".spec"
    if os.path.exists(spec):
        os.remove(spec)

    # 9) 两份等价产物只保留 .app（终端用法走 .app/Contents/MacOS/）
    removed = drop_redundant_product(app_binary, bare)
    if removed:
        print("产物去重：已移除与 .app 内容重复的裸产物（%.1f MB）"
              % (removed / (1024 * 1024)))
        print("          终端批处理请用: %s/Contents/MacOS/%s --batch ..."
              % (APP, APP))

    print("\n打包成功!")
    print("产物: %s (%.1f MB)" % (os.path.abspath(product), _size_mb(product)))
    _print_arch(binary)
    if product == app_dir:
        print("使用: 双击 .app 启动；或在终端运行")
        print("      %s/Contents/MacOS/%s --batch 文件1.xlsx 文件2.xlsx"
              % (APP, APP))
        print("分发: 发给他人时请右键 .app → 压缩，对方解压后同样右键→打开")
    else:
        print("使用: ./%s --batch 文件1.xlsx 文件2.xlsx" % APP)
    if not forced_ok:
        print("警告：无法强制 x86_64（当前 Python 不支持该目标架构），")
        print("      已按本机架构打包。要产出 Intel 版请在 Intel Mac 上重跑，")
        print("      或安装 universal2/x86_64 的 Python 3.12 后重跑。")
    return True


if __name__ == "__main__":
    if not build_exe():
        sys.exit(1)
    print("\n完成。")
