#!/usr/bin/env python3
"""
独立打包脚本
创建完全自包含的可执行文件，包含所有配置文件
"""

import os
import sys
import shutil
import subprocess
import json

def create_standalone_build():
    print("====================================")
    print("   Building SeaTable Excel Generator")
    print("====================================")
    
    # 1. Install dependencies
    print("1. Installing dependencies...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl", "python-dotenv", "seatable-api"], check=True)
    
    # 2. Clean old files
    print("\n2. Cleaning previous build files...")
    for folder in ["dist", "build"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
    
    for file in os.listdir("."):
        if file.endswith(".spec"):
            os.remove(file)
    
    # 3. Collect config files
    json_files = [f for f in os.listdir(".") if f.endswith(".json")]
    print(f"\n3. Found config files: {', '.join(json_files)}")
    
    # 4. 构建PyInstaller命令
    cmd = [
        "pyinstaller",
        "--onefile",
        "--console",
        "--name", "seatable-excel-generator",
        "--noupx",  # 禁用UPX压缩，避免DLL加载问题
        "--clean",  # 清理缓存
        "--hidden-import", "seatable_api",
        "--hidden-import", "openpyxl",
        "--hidden-import", "dotenv",
        "--hidden-import", "json",
        "--hidden-import", "datetime"
    ]
    
    # Windows特定选项
    if sys.platform.startswith("win"):
        cmd.extend([
            "--collect-all", "seatable_api",  # 收集所有seatable_api依赖
            "--collect-all", "openpyxl",      # 收集所有openpyxl依赖
            "--noconsole" if "--noconsole" in sys.argv else "--console"
        ])
    
    # 添加JSON配置文件
    json_files = [f for f in os.listdir(".") if f.endswith(".json")]
    for json_file in json_files:
        cmd.extend(["--add-data", f"{json_file}:."])
    
    # 添加.env文件（如果存在）
    if os.path.exists(".env"):
        cmd.extend(["--add-data", ".env:."])
    
    # 添加主文件
    cmd.append("main-pro.py")
    
    print(f"\n4. Executing build command...")
    print(f"Command: {' '.join(cmd)}")
    
    try:
        subprocess.run(cmd, check=True)
        print("\n[OK] Build successful!")
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] Build failed: {e}")
        return False
    
    # 5. Create deployment package
    print("\n5. Creating deployment package...")
    
    # Create deployment directory
    deploy_dir = "seatable-excel-generator-deploy"
    if os.path.exists(deploy_dir):
        shutil.rmtree(deploy_dir)
    os.makedirs(deploy_dir)
    
    # Copy executable file
    exe_name = "seatable-excel-generator.exe" if sys.platform.startswith("win") else "seatable-excel-generator"
    src_exe = os.path.join("dist", exe_name)
    dst_exe = os.path.join(deploy_dir, exe_name)
    
    if os.path.exists(src_exe):
        shutil.copy2(src_exe, dst_exe)
        print(f"[OK] Copied executable: {exe_name}")
    else:
        print(f"[ERROR] Executable not found: {src_exe}")
        return False
    
    # Copy JSON config files
    json_files = [f for f in os.listdir(".") if f.endswith(".json")]
    for json_file in json_files:
        shutil.copy2(json_file, deploy_dir)
        print(f"[OK] Copied config file: {json_file}")
    
    # Copy .env file
    if os.path.exists(".env"):
        shutil.copy2(".env", os.path.join(deploy_dir, ".env.example"))
        print("[OK] Copied .env file as .env.example")
    
    # Copy documentation
    if os.path.exists("README.md"):
        shutil.copy2("README.md", deploy_dir)
    
    # 创建使用说明
    readme_content = """# SeaTable Excel 生成器部署包

## 使用步骤：

1. 配置环境变量（推荐）：
   cp .env .env.example
   编辑 .env 文件，填入你的SeaTable Token

2. 运行工具：
   # Windows:
   seatable-excel-generator.exe
   
   # Linux/macOS:
   ./seatable-excel-generator

## 配置方式：

### .env文件配置（推荐）
1. 复制 .env 为 .env.example（如果需要）
2. 编辑 .env 文件，填入配置信息：
   - SEATABLE_SERVER_URL=你的SeaTable服务器地址
   - SEATABLE_API_TOKEN_MEMO=目标跟踪配置专用Token
   - SEATABLE_API_TOKEN_REWARD=奖励核算配置专用Token
3. 直接运行: ./seatable-excel-generator

## 配置文件说明：

- memo2025.json: 2025年目标跟踪数据配置
- reward2025.json: 2025年奖励核算数据配置

每个配置文件支持：
- 独立的SeaTable API Token配置（通过.env文件中的变量引用）
- 多个Excel文件生成配置
- 目录别名定义
- 文件合并功能配置

## 功能特性：

- 支持多配置文件自动发现
- 支持每个配置文件独立的API Token
- 智能菜单系统，选择配置文件和生成项
- 从SeaTable生成Excel文件
- 自动格式化日期和数字
- 智能处理百分比数值
- 使用SUBTOTAL函数计算合计
- 支持文件合并功能
- 跨平台支持（Windows, Linux, macOS）

## 注意事项：

- 确保网络能访问SeaTable服务
- 确保API Token有相应的表格权限
- 输出目录会自动创建（如果不存在）
- 支持.xlsx格式Excel文件
"""
    
    with open(os.path.join(deploy_dir, "USAGE.txt"), "w", encoding="utf-8") as f:
        f.write(readme_content)
    
    print("\n====================================")
    print("[SUCCESS] SeaTable Excel Generator package created successfully!")
    print(f"Package location: {deploy_dir}/")
    print(f"Executable: {deploy_dir}/{exe_name}")
    print("Share the entire folder with your team")
    print("====================================")
    
    return True

if __name__ == "__main__":
    success = create_standalone_build()
    if not success:
        sys.exit(1)