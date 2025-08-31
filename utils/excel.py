import pandas as pd
import os

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
        
        # 读取Excel文件
        print(f"正在读取文件: {file_path}")
        print(f"工作表: {sheet_name}")
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
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
            filtered_df.to_excel(output_file, index=False)
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
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
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