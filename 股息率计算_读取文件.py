import openpyxl
from openpyxl.utils import get_column_letter

def main():
    file_path = r'C:\Users\hspcadmin\Desktop\股息目标测算表_年度20W.xlsx'
    
    # 加载工作簿（保持公式，不计算值）
    try:
        wb = openpyxl.load_workbook(file_path, data_only=False)
    except FileNotFoundError:
        print(f"错误：文件未找到 - {file_path}")
        return
    except Exception as e:
        print(f"加载工作簿出错：{e}")
        return

    ws = wb.active

    # 获取标题行（第1行）的列索引
    headers = {}
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        if cell.value:
            headers[cell.value] = col

    # 检查必要列
    required = ['股数', '收盘价(元/股)', '成本元/股', '近一年股息率',
                '当前市值/元', '持仓总成本(元)', '近一年预计分红(元)', '备注']
    for name in required:
        if name not in headers:
            print(f"错误：缺少必要列 '{name}'")
            return

    # 可选列：20w目前完成度
    col_completion = '20w目前完成度'
    has_completion = col_completion in headers

    # 获取各列的字母
    col_shares = get_column_letter(headers['股数'])
    col_price = get_column_letter(headers['收盘价(元/股)'])
    col_cost = get_column_letter(headers['成本元/股'])
    col_div_yield = get_column_letter(headers['近一年股息率'])
    col_market = get_column_letter(headers['当前市值/元'])
    col_total_cost = get_column_letter(headers['持仓总成本(元)'])
    col_dividend = get_column_letter(headers['近一年预计分红(元)'])
    col_remark = get_column_letter(headers['备注'])

    # 遍历行，记录数据行范围和汇总行
    data_rows = []  # 数据行号列表
    summary_row = None
    for row in range(2, ws.max_row + 1):
        cell_a = ws.cell(row=row, column=1).value
        if cell_a and '汇总' in str(cell_a):
            summary_row = row
        else:
            # 检查是否有股数等数据（非空即视为数据行）
            if ws.cell(row=row, column=headers['股数']).value is not None:
                data_rows.append(row)

    if not data_rows:
        print("警告：未找到任何数据行，请检查文件内容。")
        return

    # 为每个数据行写入公式
    for row in data_rows:
        # 当前市值/元 = 股数 * 收盘价
        formula_market = f"={col_shares}{row}*{col_price}{row}"
        ws.cell(row=row, column=headers['当前市值/元'], value=formula_market)

        # 持仓总成本(元) = 股数 * 成本元/股
        formula_cost = f"={col_shares}{row}*{col_cost}{row}"
        ws.cell(row=row, column=headers['持仓总成本(元)'], value=formula_cost)

        # 近一年预计分红(元) = 股数 * 收盘价 * 近一年股息率
        # 注意股息率可能为百分比格式，直接用乘法
        formula_dividend = f"={col_shares}{row}*{col_price}{row}*{col_div_yield}{row}"
        ws.cell(row=row, column=headers['近一年预计分红(元)'], value=formula_dividend)

    # 处理汇总行
    if summary_row is not None:
        # 构建动态求和公式：从第2行到汇总行上一行
        first_data_row = min(data_rows) if data_rows else 2
        last_data_row = summary_row - 1

        # 当前市值合计
        if last_data_row >= first_data_row:
            sum_market = f"=SUM({col_market}{first_data_row}:{col_market}{last_data_row})"
            ws.cell(row=summary_row, column=headers['当前市值/元'], value=sum_market)

            # 持仓总成本合计
            sum_cost = f"=SUM({col_total_cost}{first_data_row}:{col_total_cost}{last_data_row})"
            ws.cell(row=summary_row, column=headers['持仓总成本(元)'], value=sum_cost)

            # 近一年预计分红合计
            sum_dividend = f"=SUM({col_dividend}{first_data_row}:{col_dividend}{last_data_row})"
            ws.cell(row=summary_row, column=headers['近一年预计分红(元)'], value=sum_dividend)

            # 20w目前完成度（如果存在）
            if has_completion:
                # 完成度 = 分红合计 / 200000
                completion_formula = f"={col_dividend}{summary_row}/200000"
                ws.cell(row=summary_row, column=headers[col_completion], value=completion_formula)

            # 总股息率写入备注列（文本公式）
            # 备注 = "总股息率: " & TEXT(分红合计/成本合计, "0.00%")
            remark_formula = (f'="总股息率: "&TEXT(IFERROR({col_dividend}{summary_row}/{col_total_cost}{summary_row},0),"0.00%")')
            ws.cell(row=summary_row, column=headers['备注'], value=remark_formula)

            print(f"汇总行公式已更新，总股息率公式：{remark_formula}")
        else:
            print("警告：没有数据行，无法写入汇总公式。")
    else:
        print("警告：未找到包含“汇总”的行，汇总数据未写入。")

    # 保存原文件
    try:
        wb.save(file_path)
        print(f"文件已成功更新（公式写入）：{file_path}")
    except Exception as e:
        print(f"保存文件时出错：{e}")

if __name__ == "__main__":
    main()