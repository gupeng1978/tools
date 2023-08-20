from openpyxl import Workbook


def merge_cells_for_records(worksheet, excel_row_start, excel_row_end, sort_key_row_index):
    
    # 遍历每一列
    cur_allowed_combined = [(excel_row_start, excel_row_end)]
    next_allowd_combined = None
    
    def try_merged(start_row, end_row, column) :
        print(cur_allowed_combined)
        worksheet.merge_cells(
            start_row=start_row, start_column=column,
            end_row=end_row, end_column=column
        )
        pass
    
    for col in worksheet.iter_cols(min_row=excel_row_start, max_row=excel_row_end):
        
        # 只合并sort_key_row_index所在的列
        if col[0].column - 1 not in sort_key_row_index:
            continue
        
        col_number = col[0].column
        combine_start = excel_row_start
        previous_value = None
        # 遍历列中的每一行
        for row_index, cell in enumerate(col, start=excel_row_start):
            if previous_value is None:
                previous_value = cell.value
                continue
            
            # 如果当前值与前一个值相同，继续遍历
            if cell.value == previous_value:
                continue
            
            # 如果当前值与前一个值不同，检查是否需要合并
            if row_index - combine_start > 1:
                try_merged(combine_start, row_index - 1, cell.column)
                # worksheet.merge_cells(
                #     start_row=combine_start, start_column=cell.column,
                #     end_row=row_index - 1, end_column=cell.column
                # )
            
            # 更新合并的起始行和前一个值
            combine_start = row_index
            previous_value = cell.value
        
        # 检查最后一组单元格是否需要合并
        if excel_row_end - combine_start > 0:
            try_merged(combine_start, excel_row_end, cell.column)
            # worksheet.merge_cells(
            #     start_row=combine_start, start_column=cell.column,
            #     end_row=excel_row_end, end_column=cell.column
            # )

# 示例
workbook = Workbook()
worksheet = workbook.active
worksheet.append(['a', 'b', 'b', 'c'])
worksheet.append(['a', 'b', 'c', 'c'])
worksheet.append(['a', 'x', 'c', 'c'])
worksheet.append(['a', 'x', 'c', 'd'])
worksheet.append(['a', 'x', 'c', 'd'])

merge_cells_for_records(worksheet, 1, 4, [0,1,2])

 # 输出表格数据到控制台
for row in worksheet.iter_rows(values_only=True):
    print(row)
workbook.save("output3.xlsx")
