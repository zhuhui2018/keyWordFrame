#encoding:utf-8

import configparser

class Config(object):
    def __init__(self,config_file_path):
        self.config_file_path = config_file_path
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file_path)

    def get_all_sections(self):
        return self.config.sections()

    def get_option(self,section_name,option_name):
        value = self.config.get(section_name,option_name)
        return value

    def all_section_items(self,section_name):
        items = self.config.items(section=section_name)
        return dict(items)

if __name__ == "__main__":
    config = Config(r"C:\Users\yanpeng\PycharmProjects\keyword_drivenn\testdata\ObjectDeposit.ini")
    print(config.get_all_sections())
    print(config.get_option("baidu","SearchPage.InputBox"))
    print(config.all_section_items("baidu")['searchpage.inputbox'])
