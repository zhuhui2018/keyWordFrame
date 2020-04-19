#encoding:utf-8

from selenium.webdriver.support.ui import WebDriverWait
from Util.GetConfig import *
import configparser
from selenium import webdriver
import os
from Config.ProjVar import *

class ObjectMap(object):
    def __init__(self,config_file_path):
        self.config_file_path = config_file_path
        self.cf = Config(self.config_file_path)

    def get_locatemethod_and_locateexpression(self,webSiteName,elementName):
        locators = self.cf.get_option(webSiteName,elementName)
        return locators

    def getElementObject(self,driver,webSiteName,elementName):
        try:
            #创建一个读取配置文件的实例
            #cf = configparser.ConfigParser()
            #将配置文件加载到内存
            #cf.read(self.config_file_path)
            #根据section和option获取配置文件页面元素的定位方式及
            #定位表达式组成的字符串，并使用“>”分割
            locators=self.cf.get_option(webSiteName,elementName).split('>')
            locatorMethod = locators[0]
            locatorExpression = locators[1]
            print(locatorMethod,locatorExpression)
            element = WebDriverWait(driver,10).until(lambda x:\
                        x.find_element(locatorMethod,locatorExpression))
        except Exception as e:
            raise e
        else:
            return element

if __name__ =="__main__":
    #driver = webdriver.Chrome(executable_path="c:\\chromedriver.exe")
    #url = "http://www.baidu.com"
    #driver.get(url)
    objmap = ObjectMap(object_map_file_path)
    #print(objmap.getElementObject(driver,"baidu","SearchPage.InputBox"))
    a,b = objmap.get_locatemethod_and_locateexpression("baidu","SearchPage.InputBox").split(">")
    print(a,b)