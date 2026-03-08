import akshare as ak
import pandas as pd

def fetch_all_stocks_to_excel(output_file='所有A股股票代码名称.xlsx'):
    """
    获取所有A股股票的中文名称和股票代码，并保存到Excel文件。
    """
    try:
        # 获取A股股票代码和名称
        df = ak.stock_info_a_code_name()
        print(f"成功获取 {len(df)} 只A股股票信息")
        
        # 保存到Excel（使用openpyxl引擎）
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"数据已保存到 {output_file}")
    except Exception as e:
        print(f"获取数据失败: {e}")

if __name__ == "__main__":
    # 可自定义输出路径，例如 r'C:\Users\你的用户名\Desktop\股票代码表.xlsx'
    fetch_all_stocks_to_excel()