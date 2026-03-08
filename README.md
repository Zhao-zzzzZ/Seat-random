# 座位分配系统

## 项目简介
本项目用于根据学生名单和特殊安排，自动分配座位，适用于教室、考场等场景。

## 主要文件说明
- `src/座位分配.py`：主程序源码，负责座位分配逻辑。
- `座位分配.py`：兼容启动入口，转发到 `src/座位分配.py`。
- `data/学生名单.json`：包含所有学生的基本信息。
- `data/特殊安排.json`：记录需要特殊安排的学生及其座位要求。
- `data/配置.json`：系统配置文件。
- `scripts/build.py`：用于打包或构建项目。
- `requirements.txt`：项目依赖库列表。
- `scripts/version.txt`：版本信息。

## 使用方法
1. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
2. 运行主程序：
   ```bash
   python 座位分配.py
   ```
   或：
   ```bash
   python src/座位分配.py
   ```
3. 根据提示修改 `data/` 目录中的配置和名单文件。

## 打包说明
如需生成可执行文件，可使用 `python scripts/build.py` 或参考 `scripts/座位分配系统.spec`。
打包完成后，程序会在 `dist/` 下生成 exe，并保留外部 `data/` 目录供读取和写入。

## 目录结构
```
├── data/
│   ├── 学生名单.json
│   ├── 特殊安排.json
│   └── 配置.json
├── requirements.txt
├── src/
│   └── 座位分配.py
├── scripts/
│   ├── build.py
│   ├── version.txt
│   └── 座位分配系统.spec
├── 座位分配.py
├── README.md
├── build/
└── dist/
```
