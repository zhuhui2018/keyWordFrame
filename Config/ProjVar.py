#encoding:utf-8

import os

proj_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print(proj_path)

firefox_driver_path="c:\\geckoDriver.exe"
Ie_driver_path="c:\\IEDriverServer.exe"
Chrome_driver_path = "c:\\chromedriver.exe"

test_data_file = os.path.join(proj_path,"testdata",
                              "测试用例.xlsx")
test_case_sheet = "测试用例"

object_map_file_path = os.path.join(proj_path,"testdata","ObjectDeposit.ini")

testcase_is_executed_col_no = 4
testcase_executed_time_col_no = 5
testcase_executed_result_col_no = 6
test_step_executed_time_col_no = 6
test_step_executed_result_col_no = 7
test_step_executed_exception_info_col_no = 8
test_step_executed_capture_pic_path_col_no = 9



if __name__ == "__main__":
    print(test_data_file)
    print(object_map_file_path)