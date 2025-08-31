#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据筛选脚本
用于从wenti.xlsx文件中筛选符合条件的数据
"""

import sys
import os

# 添加utils目录到路径
sys.path.append(os.path.join(os.path.dirname(__file__), 'utils'))

from excel import filter_excel_data, analyze_data_distribution

def main():
    print("Excel数据筛选工具")
    print("="*40)
    
    # 检查文件是否存在
    excel_file = "wenti.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"错误: 文件 '{excel_file}' 不存在")
        print("请确保文件在当前目录下")
        return
    
    # 先分析数据分布
    print("1. 分析数据分布...")
    analyze_data_distribution(excel_file)
    
    print("\n" + "="*40 + "\n")
    
    # 执行筛选
    print("2. 执行数据筛选...")
    print("筛选条件:")
    print("  - 处理状态 = '定位分发'")
    print("  - 领域责任人 = '闵赛'")
    print()
    
    result = filter_excel_data(excel_file)
    
    if result is not None and len(result) > 0:
        print(f"\n✅ 成功筛选出 {len(result)} 条记录")
        print("结果已保存到 'filtered_data.xlsx'")
    else:
        print("\n❌ 没有找到符合条件的数据")

if __name__ == "__main__":
    main()
