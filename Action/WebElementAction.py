#encoding:utf-8

from selenium import webdriver
from Config.ProjVar import *
import time
from Util.Dir import *
from Util.GetTime import *
from Util.WaitUtil import *
from Action.ResetTestCaseResult import *
from Util.Excel import *
from Util.ObjectMap import *

driver=None

def open_browser(browser_name):
    global driver

    if browser_name == "ie":
        driver= webdriver.Ie(executable_path=Ie_driver_path)
    elif browser_name == "firefox":
        driver = webdriver.Firefox(executable_path=firefox_driver_path)
    else:
        driver = webdriver.Chrome(executable_path=Chrome_driver_path)

def visit(url):
    global driver
    try:
        driver.get(url)
    except:
        print("打开浏览器失败")

def input(locate_method,locate_expression,content):
    global driver
    try:
        element = WaitUtil(driver).visibleOfElement(locate_method,locate_expression)
        element.send_keys(content)
    except:
        traceback.print_exc()
        print("输入内容出现了异常")

def sleep(duration):
    time.sleep(int(duration))

def click(locate_method,locate_expression):
    global driver
    try:
        element = WaitUtil(driver).visibleOfElement(locate_method,locate_expression)
        element.click()
    except:
        print("点击按钮出错")

def assert_word(word):
    global driver
    assert word in driver.page_source

def quit():
    global driver
    driver.quit()

def capture_pic():
    global driver
    try:
        #截图路径
        pic_dir_path = os.path.join(proj_path,"capture_pics")
        pic_dir_path = os.path.join(pic_dir_path,make_date_dir(pic_dir_path))
        pic_path = os.path.join(pic_dir_path,get_current_time()+".png")
        result = driver.get_screenshot_as_file(pic_path)
        return pic_path
    except IOError as e:
        print(e)

def execute_test_case(test_data_file_path_and_sheet_name):
    test_data_file,sheet_name = test_data_file_path_and_sheet_name.split("||")
    clear_all_executed_info(test_data_file)
    test_data_wb = ParseExcel(test_data_file)
    test_data_wb.set_sheet_by_name(sheet_name)
    success_step_num = 0
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
        return True
    else:
        return False




if __name__ == "__main__":
    execute_test_case(r"D:\keyword_drivenn\testdata\测试用例.xlsx||搜狗")