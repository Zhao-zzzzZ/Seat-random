# 座位分配系统

## 项目简介
本项目用于根据学生名单和特殊安排，自动分配座位，适用于教室、考场等场景。

## 主要文件说明
- `座位分配.py`：主程序，负责座位分配逻辑。
- `学生名单.json`：包含所有学生的基本信息。
- `特殊安排.json`：记录需要特殊安排的学生及其座位要求。
- `配置.json`：系统配置文件。
- `build.py`：用于打包或构建项目。
- `requirements.txt`：项目依赖库列表。
- `version.txt`：版本信息。

## 使用方法
1. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
2. 运行主程序：
   ```bash
   python 座位分配.py
   ```
3. 根据提示输入或修改相关配置和名单文件。

## 打包说明
如需生成可执行文件，可使用 `build.py` 或参考 `座位分配系统.spec`。

## 目录结构
```
├── build.py
├── requirements.txt
├── version.txt
├── 学生名单.json
├── 座位分配.py
├── 座位分配系统.spec
├── 特殊安排.json
├── 配置.json
└── build/座位分配系统/
```