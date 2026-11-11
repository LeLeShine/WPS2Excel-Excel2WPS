#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS Excel修复工具 - EXE打包脚本
"""

import os
import subprocess
import sys
import shutil

def build_exe():
    """打包成EXE文件"""
    print("=== WPS Excel修复工具 EXE打包 ===")
    
    # 检查PyInstaller是否已安装
    try:
        import PyInstaller
        print(f"检测到PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("未检测到PyInstaller，正在安装...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("PyInstaller安装完成")
    
    # 清理旧的构建文件
    print("\n清理旧文件...")
    for folder in ["build", "dist", "__pycache__"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"删除 {folder}")
    
    # 检查主程序文件
    if not os.path.exists("wps_repair_standalone.py"):
        print("缺少必要文件: wps_repair_standalone.py")
        return False
    print("检测到: wps_repair_standalone.py")
    
    # 图标为可选资源，存在时才打包
    icon_file = "assset/icon.ico"
    has_icon = os.path.exists(icon_file)
    if has_icon:
        print(f"检测到图标: {icon_file}")
    else:
        print("未找到 assset/icon.ico，将使用系统默认图标")
    
    # PyInstaller命令参数（用 python -m PyInstaller 调用，避免依赖 PATH）
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # 打包成单个EXE文件
        "--windowed",                   # 无控制台窗口
        "--noconfirm",                  # 覆盖旧文件不提示
        "--name=WPS_Excel_Repair_Tool", # EXE文件名
        "--distpath=.",                 # 输出到当前目录
    ]
    if has_icon:
        cmd += [f"--icon={icon_file}", f"--add-data={icon_file};assset"]
    cmd.append("wps_repair_standalone.py")  # 主程序文件
    
    print(f"\n开始打包...")
    print(f"执行命令: {' '.join(cmd[1:])}")
    
    try:
        # 执行打包命令（按 UTF-8 解码输出，系统为 GBK 代码页时也不至于崩溃）
        result = subprocess.run(cmd, check=True, capture_output=True,
                                text=True, encoding='utf-8', errors='replace')
        
        print("打包成功!")
        
        # 检查生成的EXE文件
        exe_path = "WPS_Excel_Repair_Tool.exe"
        if os.path.exists(exe_path):
            size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
            print(f"生成文件: {exe_path} ({size:.1f} MB)")
            
            # 清理构建文件
            print("\n清理构建文件...")
            for folder in ["build"]:
                if os.path.exists(folder):
                    shutil.rmtree(folder)
                    print(f"删除 {folder}")
            
            # 删除spec文件
            spec_file = "WPS_Excel_Repair_Tool.spec"
            if os.path.exists(spec_file):
                os.remove(spec_file)
                print(f"删除 {spec_file}")
            
            print(f"\n打包完成!")
            print(f"EXE文件位置: {os.path.abspath(exe_path)}")
            print(f"使用方法: 将.xlsx文件拖拽到EXE图标上即可修复")
            return True
        else:
            print("EXE文件生成失败")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        if e.stdout:
            print(f"输出: {e.stdout}")
        if e.stderr:
            print(f"错误: {e.stderr}")
        return False

if __name__ == "__main__":
    success = build_exe()
    if not success:
        sys.exit(1)
    
    # 仅在交互式控制台等待按键，自动化环境直接正常退出
    if sys.stdin is not None and sys.stdin.isatty():
        print("\n按任意键退出...")
        try:
            input()
        except EOFError:
            pass