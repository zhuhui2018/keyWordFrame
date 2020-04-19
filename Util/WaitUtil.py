#encoding:utf-8

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import traceback

class WaitUtil(object):
    def __init__(self,driver):
        self.driver = driver
        self.wait = WebDriverWait(self.driver,10)
        self.locate_method={
            "id":By.ID,
            "xpath":By.XPATH,
            "name":By.NAME,
            "link_text":By.LINK_TEXT,
            "partial_link_text":By.PARTIAL_LINK_TEXT
            }

    def presenceOfElement(self,locate_method,locate_expression):
        try:
            element = self.wait.until(lambda x:x.find_element(self.locate_method[locate_method],locate_expression))
            return element
        except:
            traceback.print_exc()
            raise Exception("所定位的素不存在")

    #显示在界面上才能找到
    def visibleOfElement(self,locate_method,locate_expression):
        try:
            element = self.wait.until(EC.visibility_of_element_located((self.locate_method[locate_method],
                                                                       locate_expression)))
            return element
        except:
            traceback.print_exc()
            raise Exception("所定位的元素不存在")

    def switchToframe(self,locate_method,locate_expression):
        try:
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((self.locate_method[locate_method],
                                                                              locate_expression)))
        except:
            traceback.print_exc()
            raise Exception("所定位的元素不存在")

if __name__ == "__main__":
    driver = webdriver.Firefox(executable_path="d:\\geckodriver.exe")
    wait_object = WaitUtil(driver)
    driver.get("http://mail.126.com")
    try:
        element = wait_object.switchToframe("xpath","//iframe[contains(@id,'x-URS-iframe')]")
        element.send_keys("光荣之路".decode("utf-8"))
        import time
        time.sleep(3)
    except:
        traceback.print_exc()
        print("素未找到！")