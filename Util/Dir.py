#encoding:utf-8

from Util.GetTime import *
import os
import traceback

def make_date_dir(dir_path):
    if os.path.exists(dir_path):
        current_date = get_current_date()
        path = os.path.join(dir_path,current_date)
        if not os.path.exists(path):
            os.mkdir(path)
        return current_date
    else:
        raise Exception("dir path does not exist!")



def make_time_dir(dir_path):
    if os.path.exists(dir_path):
        current_time = get_current_time()
        path = os.path.join(dir_path, current_time)
        if not os.path.exists(path):
            os.mkdir(path)
            return current_time
    else:
        raise Exception("dir path does not exist!")

if __name__ == "__main__":
    try:
        make_date_dir("e:\\testman")
    except:
        print("目录创建失败")