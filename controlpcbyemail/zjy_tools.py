# 一些常用的工具函数，解决一些常见的困扰--毛头@211225

import os

# 将工作目录切换到当前运行文件的目录
def change_work_dir_to_current():
    os.chdir(get_current_dir())


def get_current_dir():
    return os.path.abspath(os.path.dirname(__file__))


def get_father_dir():
    return os.path.dirname(get_current_dir())


def get_grand_father_dir():
    return os.path.dirname(get_father_dir())


def change_work_dir(dir_name):
    return os.chdir(dir_name)


def get_work_dir():
    return os.getcwd()


if __name__ == "zjy_tools":
    change_work_dir_to_current()
    print("引入zjy_tools包，并切换当前目录到" + get_current_dir())

if __name__ == '__main__':
    print(get_current_dir())
    print(get_father_dir())
    print(get_grand_father_dir())
    print(get_work_dir())
    change_work_dir_to_current()
    print(get_work_dir())
