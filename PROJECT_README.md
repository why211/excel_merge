# Excel 工具集

一个功能强大的 Excel 文件处理工具，支持文件合并、数据同步、智能去重等功能。

## 🌟 功能特性

### 📊 Excel 文件合并

-   **多文件合并**: 将多个 Excel 文件合并成一个文件
-   **智能列名匹配**: 自动识别相似的列名（如"学号"、"学生学号"）
-   **字段补充**: 自动补充缺失的字段
-   **智能去重**: 基于学号+姓名的去重处理
-   **备份保护**: 自动创建源文件备份

### 🔄 Excel 数据同步

-   **多源同步**: 支持单个或多个 Excel 文件作为数据源
-   **关联更新**: 基于关联字段的数据更新
-   **字段映射**: 灵活的字段映射配置
-   **增量更新**: 只更新变化的数据

### 🧠 智能处理

-   **列名清理**: 自动去除空格、特殊字符
-   **相似度匹配**: 基于字符串相似度的智能匹配
-   **常见变体识别**: 内置常见列名变体库
-   **冲突处理**: 智能处理数据冲突

## 🚀 快速开始

### 方式一：使用可执行文件（推荐）

1. 下载最新版本的 exe 文件
2. 双击运行 `Excel工具集.exe`
3. 根据提示选择功能并操作

### 方式二：从源码运行

```bash
# 克隆项目
git clone <repository-url>
cd excel-tools

# 安装依赖
pip install -r requirements.txt

# 运行程序
python excel_tool.py
```

## 📦 项目结构

```
excel-tools/
├── excel_tool.py          # 主程序入口
├── excel_merger.py        # Excel合并功能
├── excel_processor.py     # Excel处理功能
├── requirements.txt       # Python依赖
├── build_exe.py          # 打包脚本
├── deploy.py             # 自动化部署脚本
├── setup_git.py          # Git配置脚本
├── build.bat             # 快速打包（Windows）
├── deploy.bat            # 一键部署（Windows）
└── README.md             # 项目说明
```

## 🛠️ 开发环境

### 系统要求

-   Python 3.8+
-   Windows 10/11（推荐）
-   至少 2GB 可用内存

### 依赖包

-   `pandas>=2.0.0` - 数据处理
-   `openpyxl>=3.1.0` - Excel 文件读写
-   `xlrd>=2.0.0` - 旧版 Excel 支持
-   `pyinstaller>=5.13.0` - 打包工具

### 开发安装

```bash
# 创建虚拟环境
python -m venv .venv

# 激活虚拟环境
# Windows
.venv\Scripts\activate
# Linux/Mac
source .venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

## 📋 使用说明

### Excel 合并功能

1. **选择文件**: 选择要合并的 Excel 文件
2. **字段分析**: 程序自动分析所有字段
3. **字段选择**: 选择要包含的字段
4. **去重配置**: 配置去重规则
5. **执行合并**: 生成合并后的 Excel 文件

### Excel 同步功能

1. **选择数据源**: 选择数据源文件
2. **选择目标文件**: 选择要更新的目标文件
3. **字段映射**: 配置字段映射关系
4. **同步规则**: 设置同步规则
5. **执行同步**: 完成数据同步

### 去重说明

-   **完全相同记录**: 学号+姓名完全相同的记录自动合并
-   **姓名冲突**: 学号相同但姓名不同时，提供多种处理方式：
    -   保留第一条记录
    -   手动选择姓名
    -   创建多条记录
    -   跳过处理

## 🔧 构建和部署

### 快速打包

```bash
# Windows
build.bat

# 或使用Python脚本
python build_exe.py
```

### 完整部署

```bash
# Windows
deploy.bat

# 或使用Python脚本
python deploy.py
```

### Git 配置

```bash
python setup_git.py
```

## 📊 版本历史

### v1.0.0 (2025-08-08)

-   ✅ 初始版本发布
-   ✅ Excel 文件合并功能
-   ✅ Excel 数据同步功能
-   ✅ 智能列名匹配
-   ✅ 去重处理优化
-   ✅ 用户界面改进

## 🤝 贡献指南

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 👨‍💻 作者

**小王**

-   邮箱: [your-email@example.com]
-   GitHub: [your-github-profile]

## 🙏 致谢

感谢所有为这个项目做出贡献的开发者和用户！

## 📞 技术支持

如果您在使用过程中遇到问题，请：

1. 查看 [常见问题](FAQ.md)
2. 提交 [Issue](issues)
3. 联系开发者

---

⭐ 如果这个项目对您有帮助，请给我们一个星标！
