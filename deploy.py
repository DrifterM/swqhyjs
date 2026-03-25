"""
自动化部署脚本 - 替代Git命令行操作
"""

import os
import subprocess
import sys

def run_command(cmd):
    """执行命令并返回结果"""
    try:
        print(f"执行: {cmd}")
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"错误: {result.stderr}")
            return False, result.stderr
        print(f"成功: {result.stdout}")
        return True, result.stdout
    except Exception as e:
        print(f"异常: {e}")
        return False, str(e)

def main():
    # 设置工作目录
    work_dir = r"C:\Users\H3C\.copaw"
    os.chdir(work_dir)
    
    print("=" * 60)
    print("开始部署期货数据看板")
    print("=" * 60)
    
    # 第1步：检查Git状态
    success, output = run_command("git status")
    if not success:
        print("⚠️ Git未初始化或有问题，重新初始化...")
        run_command("git init")
    
    # 第2步：添加文件
    print("\n📁 添加修改的文件...")
    success, output = run_command("git add prototype_v2_expanded/")
    if not success:
        return False
    
    # 第3步：提交更改
    print("\n📝 提交更改...")
    success, output = run_command('git commit -m "Update futures dashboard with real data"')
    if not success:
        # 如果没有更改，直接推送
        print("没有需要提交的更改")
    
    # 第4步：检查远程地址
    print("\n🌐 配置远程仓库...")
    run_command("git remote remove origin")
    run_command("git remote add origin git@github.com:DrifterM/swqhyjs.git")
    
    # 第5步：推送到GitHub
    print("\n🚀 推送到GitHub...")
    success, output = run_command("git push -u origin main --force")
    if not success:
        print("使用HTTPS方式重试...")
        run_command("git remote remove origin")
        run_command("git remote add origin https://github.com/DrifterM/swqhyjs.git")
        success, output = run_command("git push -u origin main")
    
    if success:
        print("\n✅ 部署成功！")
        print("📊 访问网址: https://drifterm.github.io/swqhyjs/prototype_v2_expanded/")
    else:
        print("\n❌ 部署失败，请检查网络连接和权限")
    
    return success

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n建议手动执行以下命令：")
        print("cd C:\\Users\\H3C\\.copaw")
        print("git add prototype_v2_expanded/")
        print("git commit -m \"Update futures dashboard\"")
        print("git push origin main")
    
    # 保持控制台窗口
    input("\n按回车键退出...")