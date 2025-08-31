import pandas as pd
import os
import warnings

# 过滤Excel日期格式相关的警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def filter_excel_data(file_path="wenti.xlsx", sheet_name="DTS"):
    """
    从Excel文件中提取符合条件的数据
    
    筛选条件：
    - 处理状态 = "定位分发"
    - 领域责任人 = "闵赛"
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
    
    Returns:
        pandas.DataFrame: 符合条件的数据
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件 '{file_path}' 不存在")
        
        # 使用多种方式尝试读取Excel文件，处理各种格式问题
        print(f"正在读取文件: {file_path}")
        print(f"工作表: {sheet_name}")
        
        try:
            # 最安全的方式：使用字符串模式读取
            df = pd.read_excel(
                file_path, 
                sheet_name=sheet_name,
                engine='openpyxl',
                dtype=str,  # 全部读取为字符串，避免类型推断问题
                na_filter=False  # 不处理NA值
            )
            print("✅ 成功读取数据（安全模式）")
        except Exception as e:
            print(f"安全模式失败，尝试默认模式: {e}")
            # 备用方案：默认模式
            df = pd.read_excel(
                file_path, 
                sheet_name=sheet_name,
                engine='openpyxl',
                date_format=None,
                keep_default_na=True,
                na_values=['']
            )
        
        print(f"原始数据行数: {len(df)}")
        print(f"列名: {list(df.columns)}")
        
        # 检查必要的列是否存在
        required_columns = ["处理状态", "领域责任人"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"警告: 以下列不存在: {missing_columns}")
            print("可用的列名:")
            for i, col in enumerate(df.columns, 1):
                print(f"  {i}. {col}")
            return None
        
        # 应用筛选条件
        print("\n应用筛选条件:")
        print("- 处理状态 = '定位分发'")
        print("- 领域责任人 = '闵赛'")
        
        # 同时满足两个条件
        filtered_df = df[
            (df["处理状态"] == "定位分发") & 
            (df["领域责任人"] == "闵赛")
        ]
        
        print(f"\n筛选后的数据行数: {len(filtered_df)}")
        
        if len(filtered_df) > 0:
            print("\n筛选结果预览:")
            print(filtered_df.head())
            
            # 保存筛选结果到新文件
            output_file = "filtered_data.xlsx"
            filtered_df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"\n筛选结果已保存到: {output_file}")
        else:
            print("没有找到符合条件的数据")
        
        return filtered_df
        
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")
        return None

def analyze_data_distribution(file_path="wenti.xlsx", sheet_name="DTS"):
    """
    分析数据分布，帮助了解数据结构
    """
    try:
        # 使用相同的安全读取方式
        try:
            df = pd.read_excel(
                file_path, 
                sheet_name=sheet_name,
                engine='openpyxl',
                dtype=str,
                na_filter=False
            )
        except Exception:
            df = pd.read_excel(
                file_path, 
                sheet_name=sheet_name,
                engine='openpyxl'
            )
        
        print("=== 数据分析 ===")
        print(f"总行数: {len(df)}")
        print(f"总列数: {len(df.columns)}")
        
        # 分析"处理状态"列的分布
        if "处理状态" in df.columns:
            print("\n'处理状态'列的值分布:")
            status_counts = df["处理状态"].value_counts()
            for status, count in status_counts.items():
                print(f"  {status}: {count}")
        
        # 分析"领域责任人"列的分布
        if "领域责任人" in df.columns:
            print("\n'领域责任人'列的值分布:")
            owner_counts = df["领域责任人"].value_counts()
            for owner, count in owner_counts.items():
                print(f"  {owner}: {count}")
        
        return df
        
    except Exception as e:
        print(f"分析过程中出现错误: {str(e)}")
        return None

if __name__ == "__main__":
    # 分析数据分布
    print("首先分析数据分布...")
    analyze_data_distribution()
    
    print("\n" + "="*50 + "\n")
    
    # 执行筛选
    print("执行数据筛选...")
    result = filter_excel_data()
    
    if result is not None and len(result) > 0:
        print("\n=== 筛选成功 ===")
        print(f"共找到 {len(result)} 条符合条件的记录")
    else:
        print("\n=== 未找到符合条件的数据 ===")
