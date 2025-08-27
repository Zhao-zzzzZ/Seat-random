import PyInstaller.__main__
import os
import shutil

def 清理构建文件():
    """清理之前的构建文件"""
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("座位分配.spec"):
        os.remove("座位分配.spec")

def 打包程序():
    """使用PyInstaller打包程序"""
    # 清理之前的构建文件
    清理构建文件()
    
    # 定义打包参数
    PyInstaller.__main__.run([
        '座位分配.py',  # 主程序文件
        '--name=座位分配系统',  # 生成的exe名称
        '--windowed',  # 使用GUI模式
        '--onefile',  # 打包成单个exe文件
        '--add-data=特殊安排.json;.',  # 添加数据文件
        '--add-data=配置.json;.',  # 添加配置文件
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不询问确认
        '--noupx',  # 禁用UPX压缩，减少误报
        '--uac-admin',  # 请求管理员权限
        '--version-file=version.txt',  # 添加版本信息
    ])

if __name__ == "__main__":
    打包程序() 