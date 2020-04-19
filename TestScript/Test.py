#encoding:utf-8

from Config.ProjVar import *
from Util.Excel import *
from Action.WebElementAction import *
from Util.ObjectMap import *
from Action.ResetTestCaseResult import *

print(test_data_file)
def execute_test_cases(test_data_file):
    clear_all_executed_info(test_data_file)
    test_data_wb = ParseExcel(test_data_file)
    test_data_wb.set_sheet_by_name(test_case_sheet)
    col_cells = test_data_wb.get_col(4)
    success_step_num = 0
    for id,i in enumerate(range(1,len(col_cells))):
        #print(id+2,col_cells[i].value)
        if col_cells[i].value.lower() == "y":
            sheet_name=test_data_wb.get_cell_value(id+2,3)
            test_data_wb.set_sheet_by_name(sheet_name)
            max_row = test_data_wb.get_max_row()
            #print("max_row",max_row)
            for i in range(2,max_row+1):
                step_row = test_data_wb.get_row(i)
                action = step_row[1].value
                locate_method = step_row[2].value
                locate_expression = step_row[3].value
                if action:action = action.strip()
                if locate_method: locate_method = locate_method.strip()
                if locate_expression: locate_expression = locate_expression.strip()
                if locate_expression is not None and "Page." in locate_expression:
                    locate_method,locate_expression = ObjectMap(object_map_file_path).get_locatemethod_and_locateexpression \
                        (locate_method,locate_expression).split(">")
                    print("locate_method,locate_expression",locate_method,locate_expression)
                value = step_row[4].value
                #print(action,locate_method,locate_expression,value)
                if action is not None and locate_method is None and locate_expression \
                    is None and value is not None:
                    command = "%s('%s')" % (action,value)
                    print(command)
                elif action is not None and locate_method is not None and locate_expression \
                    is not None and value is not None:
                    command = "%s('%s','%s','%s')" %(action,locate_method,locate_expression,value)
                    print(command)
                elif action is not None and locate_method is None and locate_expression is None \
                    and value is None:
                    command = "%s()" %(action)
                    print(command)
                else:
                    command = "%s('%s','%s')" %(action,locate_method,locate_expression)
                    print(command)
                try:
                    eval(command)
                    test_data_wb.write_cell(i,7,"pass")
                    success_step_num+=1
                except AssertionError as e:
                    pic_path = capture_pic()
                    test_data_wb.write_cell(i, 7, "fail")
                    print("断言失败：\n" + traceback.format_exc())
                    test_data_wb.write_cell(i, 9, pic_path)
                    test_data_wb.write_cell(i, 8, traceback.format_exc())
                except Exception as e:
                    pic_path = capture_pic()
                    test_data_wb.write_cell(i,9,pic_path)
                    print(command + "\n" + e.message + traceback.format_exc())
                    test_data_wb.write_cell(i, 7, "fail")
                    test_data_wb.write_cell(i, 8, traceback.format_exc())
                test_data_wb.write_cell(i,6,get_current_DateTime())
            if success_step_num == max_row -1:
                test_data_wb.set_sheet_by_name(test_case_sheet)
                test_data_wb.write_cell(id+2,6,"pass")
            else:
                test_data_wb.set_sheet_by_name(test_case_sheet)
                test_data_wb.write_cell(id + 2, 6, "fail")
            test_data_wb.write_cell(id+2,5,get_current_DateTime())

if __name__ == "__main__":
    execute_test_cases(test_data_file)

