import random
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
import json
import os
import sys
import datetime
try:
    import openpyxl
    from openpyxl.styles import Alignment, PatternFill, Border, Side
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# 获取资源路径，兼容PyInstaller打包后的情况
def 获取资源路径(相对路径):
    """获取资源文件的绝对路径，兼容开发环境和打包后的环境"""
    if getattr(sys, 'frozen', False):
        # 如果是打包后的环境
        基础路径 = sys._MEIPASS
    else:
        # 如果是开发环境
        基础路径 = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(基础路径, 相对路径)

class 座位分配:
    """班级座位随机分配系统主类
    
    功能：
    - 提供图形界面进行座位随机分配
    - 支持特殊座位安排设置
    - 支持导出座位表到Excel
    - 支持管理员密码保护设置功能
    
    主要方法：
    - 随机分配座位(): 执行随机座位分配算法
    - 导出到Excel(): 将当前座位表导出为Excel文件
    - 设置指定排数(): 设置学生必须坐在指定排数
    - 清除设置(): 清除所有特殊安排设置
    """
    def __init__(self, root):
        """初始化座位分配系统
        
        参数:
            root: tkinter根窗口对象
        """
        self.root = root
        self.root.title("随机座位分配")
        self.root.geometry("650x480")  # 调整窗口尺寸，保持美观比例
        self.root.resizable(False, False)  # 禁止调整窗口大小，保持布局美观
        
        # 设置学生名单
        self.学生名单 = self.加载学生名单()
        
        # 座位布局：6列，左右两列各5人，中间4列各6人
        self.座位行数 = 6  # 最多的列有6行
        self.座位列数 = 6  # 总共6列
        
        # 初始化特殊安排
        self.指定排数安排 = {}  # 格式: {学生: 排数列表}
        
        # 记录当前分配结果
        self.当前分配结果 = {}  # 学生 -> (行, 列)
        
        # 用于记录点击状态
        self.第一次点击 = None  # 记录第一次点击的学生和位置
        self.第二次点击 = None  # 记录第二次点击的学生和位置
        
        # 加载特殊安排
        self.加载特殊安排()
        
        # 加载管理员密码
        self.管理员密码 = self.加载管理员密码()
        
        # 创建UI元素
        self.创建界面()
        
        # 绑定快捷键
        self.root.bind("<Control-Alt-s>", self.显示设置按钮)
    
    def 创建界面(self):
        # 创建标题
        标题框架 = tk.Frame(self.root)
        标题框架.pack(fill=tk.X, pady=(5, 5))  # 减少顶部间距
        
        标题标签 = tk.Label(标题框架, text="班级座位随机分配系统", font=("微软雅黑", 14, "bold"))
        标题标签.pack(pady=5)  # 减少标题上下间距
        
        # 创建框架
        self.主框架 = tk.Frame(self.root)
        self.主框架.pack(pady=2)  # 减少间距
        
        # 创建控制面板
        self.控制面板 = tk.Frame(self.root)
        self.控制面板.pack(pady=2)  # 减少间距
        
        # 添加按钮
        self.随机分配按钮 = tk.Button(self.控制面板, text="随机分配座位", command=self.随机分配座位, 
                              font=("微软雅黑", 10))
        self.随机分配按钮.grid(row=0, column=0, padx=5)
        
        # 添加导出Excel按钮
        if EXCEL_AVAILABLE:
            self.导出按钮 = tk.Button(self.控制面板, text="导出Excel", command=self.导出到Excel,
                                font=("微软雅黑", 10))
            self.导出按钮.grid(row=0, column=1, padx=5)
        
        # 创建设置按钮（默认隐藏）
        self.设置排数按钮 = tk.Button(self.控制面板, text="设置指定排数", command=self.设置指定排数, 
                             font=("微软雅黑", 10))
        self.清除设置按钮 = tk.Button(self.控制面板, text="清除所有设置", command=self.清除设置, 
                             font=("微软雅黑", 10))
        
        # 创建座位显示区域
        self.座位框架 = tk.Frame(self.root, bd=2, relief=tk.GROOVE)
        self.座位框架.pack(pady=5, padx=10)  # 减少上下间距
        
        # 添加讲台标识（移到上方）
        讲台标签 = tk.Label(self.座位框架, text="讲台", font=("微软雅黑", 10, "bold"), 
                        relief=tk.RAISED)
        讲台标签.grid(row=0, column=2, columnspan=2, pady=3)  # 减少讲台上下间距
        
        # 初始化座位标签
        self.座位标签 = []
        for i in range(self.座位行数):
            行标签 = []
            for j in range(self.座位列数):
                # 最左边和最右边的列只有5个座位，中间空出第6个
                if (j == 0 or j == 5) and i == 5:
                    标签 = tk.Label(self.座位框架, text="", width=9, height=2)
                else:
                    标签 = tk.Label(self.座位框架, text="空座位", width=9, height=2,
                             relief="solid", borderwidth=1, font=("微软雅黑", 9))
                    标签.grid(row=i+1, column=j, padx=2, pady=1)  # 减少座位之间的间距
                    标签.bind("<Button-1>", lambda e, row=i, col=j: self.处理座位点击(row, col))
                行标签.append(标签)
            self.座位标签.append(行标签)
        
        # 添加方向标识
        窗户标签 = tk.Label(self.座位框架, text="窗户", font=("微软雅黑", 10))
        窗户标签.grid(row=self.座位行数+1, column=0, pady=3)  # 减少间距
        
        门标签 = tk.Label(self.座位框架, text="门", font=("微软雅黑", 10))
        门标签.grid(row=self.座位行数+1, column=self.座位列数-1, pady=3)  # 减少间距
        
        # 添加底部状态栏
        状态栏 = tk.Frame(self.root, height=20)  # 减少状态栏高度
        状态栏.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.状态标签 = tk.Label(状态栏, text="", font=("微软雅黑", 9))
        self.状态标签.pack(side=tk.RIGHT, padx=10, pady=2)  # 减少状态栏内边距
    
    def 随机分配座位(self):
        """执行随机座位分配算法
        
        功能:
        - 根据座位布局和特殊安排随机分配座位
        - 优先满足有特殊安排的学生
        - 尝试最多100次分配，直到满足所有条件
        - 更新UI显示分配结果
        
        返回:
            无返回值，但会更新UI显示和当前分配结果
        """
        # 创建座位列表 - 生成所有可能的座位坐标
        座位 = []
        for i in range(self.座位行数):
            for j in range(self.座位列数):
                # 最左边和最右边的列只有5个座位(跳过第6行)
                if (j == 0 or j == 5) and i == 5:
                    continue
                座位.append((i, j))
        
        # 检查特殊安排是否可行 - 确保特殊安排不会超过可用座位数
        验证结果, 错误信息 = self.验证特殊安排()
        if not 验证结果:
            messagebox.showerror("错误", f"特殊安排无法满足: {错误信息}\n请修改后重试")
            return
        
        # 尝试多次分配 - 由于随机性，可能需要多次尝试才能满足所有特殊安排
        最大尝试次数 = 100
        for _ in range(最大尝试次数):
            # 复制学生名单和座位列表 - 每次尝试都从原始状态开始
            剩余学生 = self.学生名单.copy()
            剩余座位 = 座位.copy()
            分配结果 = {}  # 学生 -> (行, 列)
            
            # 先处理有指定排数的学生 - 确保特殊安排优先满足
            for 学生, 排数列表 in self.指定排数安排.items():
                if 学生 not in 剩余学生:
                    continue  # 学生可能已被分配
                
                # 找出指定排数的所有可用座位
                可用座位 = [座位 for 座位 in 剩余座位 if 座位[0] in 排数列表]
                if not 可用座位:
                    continue  # 没有可用座位，跳过此学生
                
                # 随机选择一个座位并分配
                座位 = random.choice(可用座位)
                分配结果[学生] = 座位
                剩余座位.remove(座位)
                剩余学生.remove(学生)
            
            # 随机分配剩余学生 - 无特殊安排的学生随机分配
            random.shuffle(剩余学生)
            for 学生 in 剩余学生:
                if not 剩余座位:
                    break  # 座位已用完
                座位 = random.choice(剩余座位)
                分配结果[学生] = 座位
                剩余座位.remove(座位)
            
            # 检查是否成功分配所有学生
            if len(分配结果) == len(self.学生名单):
                # 重置所有座位的背景色
                for i in range(self.座位行数):
                    for j in range(self.座位列数):
                        if not ((j == 0 or j == 5) and i == 5):  # 跳过角落的空座位
                            self.座位标签[i][j].config(bg="white")
                
                # 更新UI显示 - 在座位标签上显示学生姓名
                for 学生, (行, 列) in 分配结果.items():
                    self.座位标签[行][列].config(text=学生, font=("微软雅黑", 9, "bold"))
                
                # 保存当前分配结果 - 用于后续导出操作
                self.当前分配结果 = 分配结果.copy()
                
                return  # 分配成功，退出方法
        
        # 所有尝试都失败后显示错误
        messagebox.showerror("错误", "无法满足所有特殊安排，请减少限制条件后重试")
    
    def 显示设置按钮(self, event=None):
        """按下Ctrl+Alt+S时显示设置按钮"""
        密码 = simpledialog.askstring("验证", "请输入管理员密码:", show="*")
        if 密码 == self.管理员密码:
            self.设置排数按钮.grid(row=0, column=2, padx=5)
            self.清除设置按钮.grid(row=0, column=3, padx=5)
            messagebox.showinfo("成功", "设置按钮已显示")
        else:
            messagebox.showerror("错误", "密码错误")
    
    def 加载学生名单(self):
        """从学生名单文件加载学生列表
        
        功能:
        - 尝试从学生名单.json文件加载学生列表
        - 如果文件不存在或格式错误，则使用默认学生名单
        
        返回:
            list: 学生名单列表
        """
        学生名单文件 = "学生名单.json"
        默认学生名单 = [
            "敖康涵", "崔子傲", "杜欣怡", "冯禹栋", "弓子航", 
            "郭奕诚", "李秉锡", "李凡奇", "李其东", "李星哲", 
            "李一诺", "刘锦溪", "刘睿忱", "刘奕贤", "倪欣彤", 
            "牛新迪", "唐晚玉", "田煦禾", "王鼎宸", "王柳嘉", 
            "王培源", "王烁妍", "王一冉", "王子辰", "吴金航", 
            "许泽玉", "薛旭然", "杨紫斐", "张浩然", "张颀萱", 
            "张依依", "张译文", "赵鑫炜", "赵一诺"
        ]
        
        try:
            # 尝试获取打包后的路径
            try:
                学生名单路径 = 获取资源路径(学生名单文件)
            except:
                学生名单路径 = 学生名单文件
                
            if os.path.exists(学生名单路径):
                with open(学生名单路径, "r", encoding="utf-8") as f:
                    return json.load(f)
            else:
                # 如果文件不存在，创建默认学生名单文件
                with open(学生名单路径, "w", encoding="utf-8") as f:
                    json.dump(默认学生名单, f, ensure_ascii=False, indent=4)
                return 默认学生名单
        except Exception as e:
            print(f"加载学生名单时出错: {e}")
            return 默认学生名单
            
    def 加载特殊安排(self):
        """从JSON文件加载特殊安排
        
        功能:
        - 尝试从特殊安排.json文件加载特殊座位安排
        - 如果文件不存在或格式错误，则使用空字典
        - 兼容PyInstaller打包后的环境
        
        返回:
            无返回值，但会更新self.指定排数安排
        """
        特殊安排文件 = "特殊安排.json"
        
        # 尝试获取打包后的路径
        try:
            特殊安排路径 = 获取资源路径(特殊安排文件)
        except:
            特殊安排路径 = 特殊安排文件
            
        if not os.path.exists(特殊安排路径):
            # 如果文件不存在，创建一个空的特殊安排文件
            self.保存特殊安排()
            return
        
        try:
            with open(特殊安排路径, "r", encoding="utf-8") as f:
                try:
                    数据 = json.load(f)
                    
                    # 清除现有安排
                    self.指定排数安排 = {}
                    
                    # 加载指定排数安排
                    for 学生, 排数列表 in 数据.get("指定排数安排", {}).items():
                        self.指定排数安排[学生] = 排数列表
                    
                except json.JSONDecodeError as e:
                    messagebox.showerror("错误", f"特殊安排文件格式错误: {str(e)}")
                    return
        except FileNotFoundError:
            # 文件不存在，创建一个空的特殊安排文件
            self.保存特殊安排()
            return
        except PermissionError:
            messagebox.showerror("错误", "没有权限读取特殊安排文件")
            return
        except Exception as e:
            messagebox.showerror("错误", f"加载特殊安排失败: {str(e)}")
            return
    
    def 保存特殊安排(self):
        """保存特殊安排到JSON文件"""
        数据 = {
            "指定排数安排": self.指定排数安排
        }
        
        特殊安排文件 = "特殊安排.json"
        
        try:
            # 尝试获取打包后的路径
            try:
                特殊安排路径 = 获取资源路径(特殊安排文件)
            except:
                特殊安排路径 = 特殊安排文件
            
            with open(特殊安排路径, "w", encoding="utf-8") as f:
                json.dump(数据, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("错误", f"保存特殊安排失败: {str(e)}")
    
    def 设置指定排数(self):
        """设置学生坐在指定排数"""
        输入 = simpledialog.askstring("设置指定排数", "请输入学生姓名和排数(用空格分隔，排数从0开始):")
        if not 输入:
            return
        
        输入列表 = 输入.split()
        if len(输入列表) < 2:
            messagebox.showerror("错误", "请输入学生姓名和至少一个排数")
            return
        
        学生 = 输入列表[0]
        排数列表 = []
        
        # 验证学生是否在名单中
        if 学生 not in self.学生名单:
            messagebox.showerror("错误", "学生不在名单中")
            return
        
        # 解析排数
        try:
            for 排数 in 输入列表[1:]:
                排数 = int(排数)
                if 排数 < 0 or 排数 >= self.座位行数:
                    messagebox.showerror("错误", f"排数必须在0到{self.座位行数-1}之间")
                    return
                排数列表.append(排数)
        except ValueError:
            messagebox.showerror("错误", "排数必须是数字")
            return
        
        # 添加指定排数安排
        self.指定排数安排[学生] = 排数列表
        
        # 保存设置
        self.保存特殊安排()
        messagebox.showinfo("成功", f"已设置{学生}坐在第{','.join(map(str, 排数列表))}排")
    
    def 清除设置(self):
        """清除所有特殊安排"""
        # 清除指定排数安排
        self.指定排数安排 = {}
        
        # 清除座位显示
        for i in range(self.座位行数):
            for j in range(self.座位列数):
                if not ((j == 0 or j == 5) and i == 5):
                    self.座位标签[i][j].config(text="空座位", font=("微软雅黑", 9))
        
        # 保存设置
        self.保存特殊安排()
        messagebox.showinfo("成功", "已清除所有设置")
    
    def 验证特殊安排(self):
        """验证特殊座位安排是否可行
        
        功能:
        - 检查指定排数的学生数量是否超过总座位数
        - 确保特殊安排不会导致座位不足
        
        注意:
        - 当前实现仅检查学生数量是否超过总座位数
        - 不检查排数有效性(由设置指定排数方法处理)
        
        返回:
            tuple: (验证结果, 错误信息)
            - 验证结果: True表示验证通过，False表示验证失败
            - 错误信息: 验证失败时的详细错误描述
        """
        # 计算指定排数的学生数量
        指定排数学生数 = len(self.指定排数安排)
        
        # 计算总座位数 - 减去两个角落没有的座位
        总座位数 = self.座位行数 * self.座位列数 - 2
        
        # 检查学生数量是否超过可用座位数
        if 指定排数学生数 > 总座位数:
            return False, f"指定排数的学生数量({指定排数学生数})超过可用座位数({总座位数})"
        
        return True, ""  # 验证通过

    def 加载管理员密码(self):
        """从配置文件加载管理员密码"""
        try:
            with open("配置.json", "r", encoding="utf-8") as f:
                配置 = json.load(f)
                return 配置.get("管理员密码", "admin")  # 默认密码为admin
        except:
            return "admin"  # 如果配置文件不存在，使用默认密码

    def 导出到Excel(self):
        """导出当前座位表到Excel文件
        
        功能:
        - 创建包含两个工作表的Excel文件:
          1. 座位表(学生视角): 从学生角度看的座位布局
          2. 座位表(讲台视角): 从讲台角度看的座位布局(行列翻转)
        - 添加标题、讲台标识和方向标识
        - 设置单元格格式(居中、边框等)
        - 自动生成带时间戳的文件名
        
        返回:
            无返回值，但会显示导出成功或失败的提示信息
        """
        if not self.当前分配结果:
            messagebox.showerror("错误", "请先进行座位分配")
            return
            
        # 创建新的工作簿
        wb = openpyxl.Workbook()
        
        # 创建第一个工作表（正常视图）- 学生视角
        ws1 = wb.active
        ws1.title = "座位表（学生视角）"
        
        # 创建第二个工作表（翻转视图）- 讲台视角
        ws2 = wb.create_sheet("座位表（讲台视角）")
        
        # 设置列宽 - 统一所有列的宽度为15个字符
        for col in range(1, self.座位列数 + 1):
            ws1.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15
            ws2.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15
        
        # 添加标题（仅学生视角）- 合并第一行的所有列
        ws1.merge_cells(f'A1:{openpyxl.utils.get_column_letter(self.座位列数)}1')
        标题单元格 = ws1.cell(1, 1, "班级座位表")
        标题单元格.alignment = Alignment(horizontal='center', vertical='center')
        标题单元格.font = openpyxl.styles.Font(size=14, bold=True)
        
        # 添加讲台（学生视角 - 顶部）- 合并C2和D2单元格
        ws1.merge_cells(f'C2:D2')
        讲台单元格 = ws1.cell(2, 3, "讲台")
        讲台单元格.alignment = Alignment(horizontal='center', vertical='center')
        讲台单元格.font = openpyxl.styles.Font(bold=True)
        
        # 添加座位（正常视图）- 遍历所有分配结果
        for 学生, (行, 列) in self.当前分配结果.items():
            # 调整行号（Excel从1开始，且第1行是标题，第2行是讲台）
            excel行 = 行 + 3
            excel列 = 列 + 1
            
            # 设置单元格值（正常视图）- 学生姓名居中显示
            单元格 = ws1.cell(excel行, excel列, 学生)
            单元格.alignment = Alignment(horizontal='center', vertical='center')
            单元格.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # 设置单元格值（翻转视图）- 行列位置翻转
            # 计算翻转后的位置：行和列都翻转
            翻转行 = self.座位行数 - 行 - 1
            翻转列 = self.座位列数 - 列 - 1
            翻转excel行 = 翻转行 + 1  # 讲台视角不需要标题和讲台行，所以从第1行开始
            翻转excel列 = 翻转列 + 1
            
            翻转单元格 = ws2.cell(翻转excel行, 翻转excel列, 学生)
            翻转单元格.alignment = Alignment(horizontal='center', vertical='center')
            翻转单元格.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        # 添加讲台（讲台视角 - 底部，紧贴座位区域）
        ws2.merge_cells(f'C{self.座位行数 + 1}:D{self.座位行数 + 1}')
        翻转讲台单元格 = ws2.cell(self.座位行数 + 1, 3, "讲台")
        翻转讲台单元格.alignment = Alignment(horizontal='center', vertical='center')
        翻转讲台单元格.font = openpyxl.styles.Font(bold=True)
        
        # 添加方向标识（仅学生视角）- 窗户和门标识
        窗户单元格 = ws1.cell(self.座位行数 + 3, 1, "窗户")
        窗户单元格.font = openpyxl.styles.Font(color="0000FF")
        
        门单元格 = ws1.cell(self.座位行数 + 3, self.座位列数, "门")
        门单元格.font = openpyxl.styles.Font(color="0000FF")
        
        # 保存文件 - 使用当前时间生成文件名
        文件名 = f"座位表_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        文件路径 = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=文件名,
            filetypes=[("Excel文件", "*.xlsx")]
        )
        
        if 文件路径:
            try:
                wb.save(文件路径)
                messagebox.showinfo("成功", f"座位表已导出到：\n{文件路径}\n\n包含两个工作表：\n1. 座位表（学生视角）\n2. 座位表（讲台视角）")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{str(e)}")

    def 处理座位点击(self, row, col):
        """处理座位点击事件
        
        功能:
        - 记录第一次和第二次点击的座位
        - 当点击两个座位后，执行座位互换
        - 更新界面显示和状态栏提示
        
        参数:
            row: 点击的座位行号
            col: 点击的座位列号
        """
        # 如果还没有进行座位分配，直接返回
        if not self.当前分配结果:
            self.状态标签.config(text="请先进行座位分配")
            return
            
        # 获取点击的座位上的学生
        当前学生 = None
        for 学生, (学生行, 学生列) in self.当前分配结果.items():
            if 学生行 == row and 学生列 == col:
                当前学生 = 学生
                break
                
        if not 当前学生:
            self.状态标签.config(text="请点击有学生的座位")
            return
            
        # 如果是第一次点击
        if self.第一次点击 is None:
            self.第一次点击 = (当前学生, row, col)
            self.状态标签.config(text=f"已选择{当前学生}，请选择要交换的学生")
            # 高亮显示选中的座位
            self.座位标签[row][col].config(bg="lightblue")
            return
            
        # 如果是第二次点击
        if self.第二次点击 is None:
            # 如果点击的是同一个座位
            if self.第一次点击[0] == 当前学生:
                self.状态标签.config(text="请选择不同的学生进行交换")
                return
                
            self.第二次点击 = (当前学生, row, col)
            self.状态标签.config(text=f"正在交换{self.第一次点击[0]}和{当前学生}的座位")
            
            # 执行座位互换
            self.互换座位()
            
            # 重置点击状态
            self.第一次点击 = None
            self.第二次点击 = None
            
    def 互换座位(self):
        """执行座位互换操作
        
        功能:
        - 交换两个学生的座位位置
        - 更新界面显示
        - 更新当前分配结果
        """
        if not (self.第一次点击 and self.第二次点击):
            return
            
        学生1, 行1, 列1 = self.第一次点击
        学生2, 行2, 列2 = self.第二次点击
        
        # 更新当前分配结果
        self.当前分配结果[学生1] = (行2, 列2)
        self.当前分配结果[学生2] = (行1, 列1)
        
        # 更新界面显示
        self.座位标签[行1][列1].config(text=学生2)
        self.座位标签[行2][列2].config(text=学生1)
        
        # 重置所有座位的背景色
        for i in range(self.座位行数):
            for j in range(self.座位列数):
                if not ((j == 0 or j == 5) and i == 5):  # 跳过角落的空座位
                    self.座位标签[i][j].config(bg="white")
        
        # 更新状态栏
        self.状态标签.config(text=f"已成功交换{学生1}和{学生2}的座位")

if __name__ == "__main__":
    root = tk.Tk()
    app = 座位分配(root)
    root.mainloop()