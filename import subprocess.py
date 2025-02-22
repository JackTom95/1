import subprocess
import os

def pack_to_exe(script_path, output_dir=None, icon_path=None):
    """
    使用PyInstaller将Python脚本打包为单文件exe。
    
    参数:
        script_path (str): Python脚本的路径。
        output_dir (str, 可选): 输出目录，默认为脚本所在目录的dist文件夹。
        icon_path (str, 可选): 自定义图标文件路径。
    """
    # 检查脚本路径是否存在
    if not os.path.isfile(script_path):
        print("错误：脚本文件不存在！")
        return

    # 设置输出目录
    if output_dir:
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir)
        dist_path = os.path.join(output_dir, "dist")
    else:
        dist_path = os.path.join(os.path.dirname(script_path), "dist")

    # 构造PyInstaller命令
    command = [
        "pyinstaller",
        "--onefile",  # 单文件模式
        "--noconsole",  # 隐藏控制台窗口
        "--distpath", dist_path,  # 输出路径
        script_path
    ]

    # 如果提供了图标路径，添加图标参数
    if icon_path and os.path.isfile(icon_path):
        command.extend(["--icon", icon_path])

    # 执行PyInstaller命令
    try:
        print("正在打包脚本为exe...")
        subprocess.run(command, check=True)
        print(f"打包完成！生成的exe文件位于：{dist_path}")
    except subprocess.CalledProcessError as e:
        print(f"打包失败：{e}")
    except Exception as e:
        print(f"发生错误：{e}")

# 示例用法
if __name__ == "__main__":
    script_path = input("请输入Python脚本的路径：")
    output_dir = input("请输入输出目录（可选，按回车跳过）：").strip() or None
    icon_path = input("请输入图标文件路径（可选，按回车跳过）：").strip() or None

    pack_to_exe(script_path, output_dir, icon_path)