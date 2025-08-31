import pandas as pd
import numpy as np

def create_test_excel():
    """
    创建测试用的wenti.xlsx文件
    """
    # 创建测试数据
    data = {
        '问题ID': ['DTS-001', 'DTS-002', 'DTS-003', 'DTS-004', 'DTS-005', 'DTS-006', 'DTS-007', 'DTS-008'],
        '问题标题': [
            '系统登录问题',
            '数据同步异常', 
            '页面加载缓慢',
            '接口超时',
            '用户权限错误',
            '数据库连接失败',
            '文件上传问题',
            '报表生成错误'
        ],
        '处理状态': [
            '定位分发',  # 符合条件
            '已解决',
            '定位分发',  # 符合条件
            '处理中',
            '定位分发',  # 符合条件
            '已关闭',
            '定位分发',  # 符合条件
            '待处理'
        ],
        '领域责任人': [
            '闵赛',      # 符合条件
            '张三',
            '李四',
            '闵赛',      # 不符合条件（状态不对）
            '闵赛',      # 符合条件
            '王五',
            '闵赛',      # 符合条件
            '赵六'
        ],
        '优先级': ['高', '中', '低', '高', '中', '低', '高', '中'],
        '创建时间': [
            '2025-08-25',
            '2025-08-26', 
            '2025-08-27',
            '2025-08-28',
            '2025-08-29',
            '2025-08-30',
            '2025-08-31',
            '2025-08-31'
        ],
        '描述': [
            '用户无法正常登录系统',
            '数据同步出现延迟',
            '页面响应时间过长',
            'API接口调用超时',
            '用户权限配置有误',
            '无法连接到数据库',
            '文件上传功能异常',
            '月度报表生成失败'
        ]
    }
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存为Excel文件，指定工作表名为DTS
    with pd.ExcelWriter('wenti.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='DTS', index=False)
    
    print("测试文件 'wenti.xlsx' 创建成功！")
    print(f"包含 {len(df)} 条测试数据")
    
    # 显示数据预览
    print("\n数据预览:")
    print(df)
    
    # 显示符合条件的数据统计
    filtered = df[(df['处理状态'] == '定位分发') & (df['领域责任人'] == '闵赛')]
    print(f"\n符合筛选条件的数据: {len(filtered)} 条")
    print("筛选条件: 处理状态='定位分发' AND 领域责任人='闵赛'")
    
    if len(filtered) > 0:
        print("\n符合条件的数据:")
        print(filtered[['问题ID', '问题标题', '处理状态', '领域责任人']].to_string(index=False))
    
    return df

if __name__ == "__main__":
    create_test_excel()
