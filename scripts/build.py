import PyInstaller.__main__
import os
import shutil

项目根目录 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
数据目录 = os.path.join(项目根目录, "data")
输出目录 = os.path.join(项目根目录, "dist")

def 清理构建文件():
    """清理之前的构建文件"""
    构建目录 = os.path.join(项目根目录, "build")
    if os.path.exists(构建目录):
        shutil.rmtree(构建目录)
    if os.path.exists(输出目录):
        shutil.rmtree(输出目录)

def 打包程序():
    """使用 PyInstaller 打包程序，并复制外部数据目录"""
    清理构建文件()

    PyInstaller.__main__.run([
        os.path.join(项目根目录, '座位分配.py'),
        '--name=座位分配系统',
        '--windowed',
        '--onefile',
        '--clean',
        '--noconfirm',
        '--noupx',
        '--uac-admin',
        f'--version-file={os.path.join(项目根目录, "scripts", "version.txt")}',
    ])

    目标数据目录 = os.path.join(输出目录, "data")
    shutil.copytree(数据目录, 目标数据目录, dirs_exist_ok=True)

if __name__ == "__main__":
    打包程序()
