#encoding:utf-8

from Util.Excel import *
from Config.ProjVar import *
from Action.WebElementAction import *

def clear_testcase_result(test_data_file):
    try:
        test_data_wb = ParseExcel(test_data_file)
        sheets = test_data_wb.get_all_sheet_names()
        test_data_wb.set_sheet_by_name(sheets[0])
        col_cells = test_data_wb.get_col(testcase_is_executed_col_no)
        for id,i in enumerate(range(1,len(col_cells))):
            if col_cells[i].value.lower() == "y":
                test_data_wb.write_cell(id+2,testcase_executed_time_col_no,"")
                test_data_wb.write_cell(id + 2, testcase_executed_result_col_no, "")
    except:
        return False
    else:
        return True

def clear_test_step_result(test_data_file):
    try:
        test_data_wb = ParseExcel(test_data_file)
        sheets = test_data_wb.get_all_sheet_names()
        test_data_wb.set_sheet_by_name(sheets[0])
        #测试用例是否执行列
        col_cells = test_data_wb.get_col(4)
        test_step_execute_sheet_names = []
        for id,i in enumerate(range(1,len(col_cells))):
            row = test_data_wb.get_row(id+2)
            if col_cells[i].value.lower() == "y":
                test_step_execute_sheet_names.append(row[2].value)
        for sheet in sheets[1:]:
            if sheet not in test_step_execute_sheet_names:
                continue
            test_data_wb.set_sheet_by_name(sheet)
            for row_no in range(2,test_data_wb.get_max_row()+1):
                test_data_wb.write_cell(row_no, test_step_executed_time_col_no, "")
                test_data_wb.write_cell(row_no, test_step_executed_result_col_no , "")
                test_data_wb.write_cell(row_no, test_step_executed_exception_info_col_no, "")
                test_data_wb.write_cell(row_no, test_step_executed_capture_pic_path_col_no , "")
    except:
        return False
    else:
        return True

def clear_all_executed_info(test_data_file):
    clear_result1=clear_testcase_result(test_data_file)
    clear_result2=clear_test_step_result(test_data_file)
    if clear_result1 and clear_result2:
        return True
    else:
        return False


if __name__ =="__main__":
    clear_testcase_result(test_data_file)
    clear_test_step_result(test_data_file)



