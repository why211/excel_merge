import pandas as pd
import os

def create_test_data():
    """创建测试数据"""
    
    # 创建源文件数据
    source_data = {
        '教工号': ['sy001', 'sy002', 'sy003', 'sy004', 'sy005'],
        '学生姓名': ['张三', '李四', '王五', '赵六', '钱七'],
        'DEPT_NAME': ['计算机系', '数学系', '物理系', '化学系', '生物系'],
        '*教职工考核单位号': ['0101', '0102', '0103', '0104', '0105']
    }
    
    # 创建目标文件数据（包含一些不匹配的记录）
    target_data = {
        '教工号': ['sy001', 'sy002', 'sy003', 'sy006', 'sy007'],  # sy006, sy007 不在源文件中
        '*教师姓名': ['教师A', '教师B', '教师C', '教师D', '教师E'],
        '教职工考核类别码': ['A', 'B', 'C', 'D', 'E'],
        '*教职工考核日期': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04', '2024-01-05'],
        '教职工考核内容': ['内容1', '内容2', '内容3', '内容4', '内容5'],
        '*教职工考核单位号': ['', '', '', '', ''],  # 空值，需要更新
        '*考核单位名称': ['单位1', '单位2', '单位3', '单位4', '单位5'],
        '*单位考核结果码': ['R1', 'R2', 'R3', 'R4', 'R5'],
        '单位考核负责人号': ['F1', 'F2', 'F3', 'F4', 'F5'],
        '学校考核结果码': ['S1', 'S2', 'S3', 'S4', 'S5'],
        '*发起人姓名': ['发起人1', '发起人2', '发起人3', '发起人4', '发起人5'],
        '*发 起人（工号/学号）': ['G1', 'G2', 'G3', 'G4', 'G5'],
        '数据统计部门': ['部门1', '部门2', '部门3', '部门4', '部门5']
    }
    
    # 创建DataFrame
    source_df = pd.DataFrame(source_data)
    target_df = pd.DataFrame(target_data)
    
    # 保存到文件
    source_file = r"G:\wang\excel\test_source.xls"
    target_file = r"G:\wang\excel\test_target.xls"
    
    with pd.ExcelWriter(source_file, engine='openpyxl') as writer:
        source_df.to_excel(writer, index=False)
    
    with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
        target_df.to_excel(writer, index=False)
    
    print(f"✅ 测试数据已创建:")
    print(f"源文件: {source_file}")
    print(f"目标文件: {target_file}")
    print(f"\n📊 数据说明:")
    print(f"源文件记录数: {len(source_df)}")
    print(f"目标文件记录数: {len(target_df)}")
    print(f"匹配的记录: sy001, sy002, sy003 (3个)")
    print(f"不匹配的记录: sy006, sy007 (2个)")

if __name__ == "__main__":
    create_test_data() 