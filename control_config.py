# @time 2021/9/18 11:27
# @Author howell
# @File controlConfig.PY
import os
import configparser


class ReadConfig(object):
    def __init__(self, filepath=None):
        if filepath:
            self.confiscate = filepath
        else:
            # 获取conf文件中config.ini文件的路径
            self.confiscate = os.path.abspath(
                os.path.join(os.path.dirname('__file__'), os.path.pardir, 'conf', 'config.ini'))
        self.cf = configparser.ConfigParser()
        self.cf.read(self.confiscate, encoding="utf-8")

    def get(self, *args):
        """
        :param args: args[0]==section args[1]==option
        :return: key for value
        """
        return self.cf.get(args[0], args[1])


class WriteConfig(object):
    def __init__(self, filepath=None):
        if filepath:
            self.confiscate = filepath
        else:
            # 获取conf文件中config.ini文件的路径
            self.confiscate = os.path.abspath(
                os.path.join(os.path.dirname('__file__'), os.path.pardir, 'conf', 'writedConfig.ini'))
        self.cf = configparser.ConfigParser()
        self.cf.read(self.confiscate, encoding="utf-8")

    def write(self, section, option, value):
        """
        :param section: 分类
        :param option: 对应的键
        :param value: 值
        :return: 返回添加成功的结果
        """
        if not self.cf.has_section(section):
            self.cf.add_section(section)
        self.cf.set(section, option, value)
        self.cf.write(open(self.confiscate, 'r+', encoding='utf-8'))


if __name__ == '__main__':
    filepath = os.getcwd() + r"\appium_config.ini"
    # print(ReadConfig(filepath).get("run_data", "run_num"))
    WriteConfig(filepath).write("run_data", "start_num", "1108")
