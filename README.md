# Excel 文件处理工具

这是一个用于处理 Excel 文件的 Python 工具，可以合并多个 Excel 文件中的指定字段，并支持去重和字段分析功能。

## 功能特性

### 1. 字段分析功能

-   扫描指定文件夹中的所有 Excel 文件
-   统计所有文件中出现的字段（列名）总数
-   显示每个字段在哪些文件中出现
-   生成详细的字段统计报告

### 2. 可选择的去重功能

-   在输出结果时，提供选项让用户选择是否进行去重
-   如果选择去重，则按学号去重；如果不选择去重，则保留所有重复记录
-   在交互式版本中，通过用户输入来控制去重选项
-   在简化版本中，通过参数来控制去重选项

### 3. 基于字段选择的文件合并功能

-   根据用户输入的字段名，确定哪些 Excel 文件包含这些字段
-   只处理包含指定字段的文件
-   如果某个文件缺少指定字段，则跳过该文件
-   支持多个字段的选择

## 文件说明

### excel_processor.py

交互式版本，提供完整的用户交互界面：

-   支持菜单选择操作
-   交互式输入参数
-   字段分析功能
-   可选择的去重功能
-   自定义字段选择

### excel_processor_simple.py

简化版本，通过修改代码中的参数来使用：

-   预设参数，无需交互
-   适合批量处理
-   支持所有核心功能

### example_usage.py

示例使用脚本，展示各种功能的使用方法。

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 交互式使用（推荐）

```bash
python excel_processor.py
```

运行后会显示菜单：

```
=== Excel文件处理工具 ===
1. 分析字段
2. 处理文件
3. 分析字段并处理文件
```

选择操作后，程序会引导您完成后续设置。

### 2. 简化版本使用

修改 `excel_processor_simple.py` 中的参数：

```python
if __name__ == "__main__":
    # 在这里设置你的参数
    folder_path = "excel"  # 修改为excel文件夹路径
    output_filename = "学生名单.xlsx"  # 输出文件名
    required_fields = ["学号", "*学生姓名"]  # 需要合并的字段
    deduplicate = True  # 是否去重

    # 可选：先分析字段
    # analyze_excel_fields(folder_path)

    # 执行处理
    process_excel_files(folder_path, output_filename, required_fields, deduplicate)
```

然后运行：

```bash
python excel_processor_simple.py
```

### 3. 编程接口使用

```python
from excel_processor import analyze_excel_fields, process_excel_files

# 分析字段
analyze_excel_fields("excel")

# 处理文件（不去重）
process_excel_files("excel", "result.xlsx",
                   required_fields=["学号", "*学生姓名"],
                   deduplicate=False)

# 处理文件（去重）
process_excel_files("excel", "result.xlsx",
                   required_fields=["学号", "*学生姓名"],
                   deduplicate=True)
```

## 功能详解

### 字段分析功能

`analyze_excel_fields(folder_path)` 函数会：

1. 扫描指定文件夹中的所有 Excel 文件
2. 读取每个文件的列名
3. 统计每个字段出现的次数和文件列表
4. 生成详细的统计报告

输出示例：

```
=== 字段统计报告 ===
总共发现 15 个不同的字段

字段出现情况：

字段: 学号
  出现次数: 8
  出现文件: 学生名单.xlsx, 奖助贷申请数据子类-国家助学金.xls, ...

字段: *学生姓名
  出现次数: 8
  出现文件: 学生名单.xlsx, 奖助贷申请数据子类-国家助学金.xls, ...
```

### 文件选择功能

`select_files_by_fields(folder_path, required_fields)` 函数会：

1. 检查每个 Excel 文件是否包含所有必需字段
2. 返回包含所有必需字段的文件列表
3. 显示哪些文件被选中，哪些被跳过

### 去重功能

处理函数支持可选择的去重：

-   `deduplicate=True`：按学号去重（如果存在学号字段）
-   `deduplicate=False`：保留所有记录，不去重

## 参数说明

### process_excel_files 函数参数

-   `folder_path` (str): 包含 Excel 文件的文件夹路径
-   `output_filename` (str): 输出文件名，默认为"result.xlsx"
-   `required_fields` (list): 必需的字段列表，默认为["学号", "*学生姓名"]
-   `deduplicate` (bool): 是否去重，默认为 True

### analyze_excel_fields 函数参数

-   `folder_path` (str): 包含 Excel 文件的文件夹路径

### select_files_by_fields 函数参数

-   `folder_path` (str): 包含 Excel 文件的文件夹路径
-   `required_fields` (list): 必需的字段列表

## 输出文件

处理完成后会生成一个 Excel 文件，包含：

-   用户选择的字段数据
-   来源文件列（用于追踪数据来源）
-   根据去重设置处理后的数据

## 错误处理

程序包含完善的错误处理机制：

-   文件不存在检查
-   文件占用处理
-   缺少字段的警告
-   读取错误的处理
-   权限错误的处理

## 示例

运行示例脚本查看各种功能的使用方法：

```bash
python example_usage.py
```

## 注意事项

1. 确保 Excel 文件没有被其他程序打开
2. 字段名必须完全匹配（包括空格和特殊字符）
3. 去重功能需要"学号"字段存在
4. 建议先使用字段分析功能了解数据结构

## 更新日志

### v2.0

-   新增字段分析功能
-   新增可选择的去重功能
-   新增基于字段选择的文件合并功能
-   改进用户交互界面
-   添加示例使用脚本

### v1.0

-   基础 Excel 文件合并功能
-   学号和姓名字段提取
-   自动去重功能
