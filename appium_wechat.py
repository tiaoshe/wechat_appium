import random
from appium import webdriver
from appium.webdriver.common.appiumby import By
import time, os
from read_excel import ExcelUtil
from appium.webdriver.common.touch_action import TouchAction
from selenium.webdriver.support.wait import WebDriverWait
from howell_logging import logger
from control_config import ReadConfig, WriteConfig


class BaseSet(object):
    def __init__(self):
        filepath = os.getcwd() + r"\appium_config.ini"
        reader = ReadConfig(filepath)
        desired_caps = {
            "platformName": reader.get("sys_data", "platformName"),
            "platformVersion": reader.get("sys_data", "platformVersion"),
            "deviceName": reader.get("sys_data", "deviceName"),
            "appPackage": reader.get("sys_data", "appPackage"),
            "appActivity": reader.get("sys_data", "appActivity"),
            "noReset": True,
            # 不关闭app直接从当前页面获取信息
            # "dontStopAppOnReset": True,
        }
        self.driver = webdriver.Remote(reader.get("sys_data", "appiumPath"), desired_caps)
        self.driver.implicitly_wait(10)


def read_excl(begin_number, run_number):
    path = os.getcwd() + r"\data\temp_data.xlsx"
    excel_obj = ExcelUtil(path).get_company_info(begin_number, run_number)
    return excel_obj


def arrangement_data(company_list, building_name="雄飞中心", building_id="5014805079"):
    arrangement_data_list = []
    data = {}
    for company in company_list:
        computed = choice_recommend_info()
        data['building_name'] = building_name
        data['building_id'] = building_id
        data['company_name'] = company[0]
        data['customer_company_name'] = company[2]
        data['contacts_phone'] = check_phone_number(company[22])
        data['contacts_type'] = "潜在客户"
        data['competitor_package'] = computed[0]
        data['competitor_money'] = computed[1]
        data['recommend_info'] = computed[2]
        data['estimated_income'] = computed[3]
        data['marketing'] = random.choice(["到期之后跟进", "定期拜访", "互留联系方式"])
        data['summary'] = random.choice(["潜在客户可以发展", "持续跟进客户", "有望发展成客户"])
        arrangement_data_list.append(data)
        data = {}
    return arrangement_data_list


def check_phone_number(phone):
    if phone == "-":
        old = "13678"
        phone = random.randint(100000, 999999)
        return old + str(phone)
    else:
        return phone


def choice_recommend_info():
    package = random.choice(['电信', '移动'])
    old = 0
    if package == "电信":
        old = random.choice([200, 600, 1200])
    elif package == "移动":
        old = random.choice([199, 299, 599])

    com_dict = {"239": "沃商务融合5G版本", "299": "沃商务融合5G版本", "399": "沃商务融合5G版本", "599": "沃商务融合5G版本", "120": "沃快车融合套餐",
                "100": "沃快车融合套餐",
                "240": "沃快车融合套餐",
                "200": "沃快车融合套餐",
                "360": "沃快车融合套餐",
                "300": "沃快车融合套餐",
                "1500": "沃专线",
                "2400": "沃专线",
                "800": "沃动车",
                "1600": "沃动车"}
    min_temp = 5000
    data_key = ""
    for x in com_dict.keys():
        if 0 < (int(x) - old) < min_temp:
            min_temp = int(x) - old
            data_key = x
    return package, old, com_dict[data_key], data_key


class SetInfo(BaseSet):
    def __init__(self):
        super().__init__()
        self.x = self.driver.get_window_size()['width']
        self.y = self.driver.get_window_size()['height']
        filepath = os.getcwd() + r"\appium_config.ini"
        self.writer = WriteConfig(filepath)
        self.reader = ReadConfig(filepath)

    def swip_up(self, times=1):
        for i in range(times):
            time.sleep(0.5)
            self.driver.swipe(980, 1700, 987, 1242, 850)

    def find_table(self):
        # 点击汇报按钮
        report_button = self.find_element_util("//android.widget.TextView[@text='汇报']")
        report_button.click()
        time.sleep(0.5)
        # 点击三个点
        fill_on_button = self.find_element_util("com.tencent.wework:id/kuz", "ID")
        fill_on_button.click()
        time.sleep(0.5)
        # 点击进入应用
        in_button = self.find_element_util("com.tencent.wework:id/buu", "ID")
        in_button.click()
        time.sleep(0.5)
        # 点击直销楼长排放记录
        submi_page_inner = self.find_element_util(
            "//android.widget.TextView[@text='日报']/../../../android.widget.FrameLayout[last()]")
        submi_page_inner.click()
        time.sleep(0.5)

    def re_in_form(self):
        # 点击直销楼长排放记录
        submi_page_inner = self.find_element_util(
            "//android.widget.TextView[@text='日报']/../../../android.widget.FrameLayout[last()]")
        submi_page_inner.click()

    def find_element_util(self, location, ftype='Xpath'):
        try:
            if ftype == 'Xpath':
                element_obj = WebDriverWait(self.driver, 10, 0.5).until(lambda x: x.find_element(By.XPATH, location))
                return element_obj
            elif ftype == 'ID':
                element_obj = WebDriverWait(self.driver, 10, 0.5).until(lambda x: x.find_element(By.ID, location))
                return element_obj
        except Exception as e:
            self.save_img("error")
            logger.error(location)
            logger.error(e)

    def save_img(self, message):
        path = os.getcwd()
        time.sleep(0.2)
        name = "howell_" + message + "_" + str(int(time.time()))
        png_path = path + r"\img\%s.png" % name
        self.driver.get_screenshot_as_file(png_path)

    def number_key_input(self, key_value):
        key_dic = {'0': 7, '1': 8, '2': 9, '3': 10, '4': 11, '5': 12, '6': 13, '7': 14, '8': 15, '9': 16}
        for i in key_value:
            if key_dic.get(i):
                self.driver.press_keycode(key_dic[i])
        self.driver.keyevent(61)

    def select_time(self, times):
        x = self.x
        y = self.y
        x1 = float(x * random.randint(45, 60) / 100)
        y1 = float(y * random.randint(90, 95) / 100)
        y2 = float(y * random.randint(60, 70) / 100)
        x2 = x1 + random.randint(80, 100)
        for i in range(times):
            time.sleep(0.5)
            self.driver.swipe(x1, y1, x2, y2, random.randint(120, 200))
        tx = 105
        ty = 1930
        time.sleep(1)
        # TouchAction(self.driver).long_press(x=tx, y=ty, duration=200).release().perform()
        TouchAction(self.driver).press(x=tx, y=ty).release().perform()

    def select_time2(self):
        a = time.localtime(time.time())
        now_mon = a[1]
        now_day = a[2]
        ran_mon = random.randint(1, 12)
        ran_day = random.randint(1, 31)
        ran_updown_mon = 1 if (ran_mon - now_mon < 0) else 0
        ran_updown_day = 1 if (ran_day - now_day < 0) else 0
        y1, y2 = 2080, 2000
        time.sleep(1)
        for i in range(random.randint(1, 2)):
            self.driver.swipe(230, y1, 235, y2, 300)
        time.sleep(0.5)
        if ran_updown_mon == 1:
            y1, y2 = 2000, 2070
        for i in range(abs(ran_mon - now_mon)):
            self.driver.swipe(557, y1, 560, y2, 500)
        time.sleep(0.5)
        if ran_updown_day == 1:
            y1, y2 = 2000, 2070
        else:
            y1, y2 = 2070, 2000
        time.sleep(0.5)
        for i in range(abs(ran_day - now_day)):
            self.driver.swipe(900, y1, 905, y2, 400)
            time.sleep(0.2)
        TouchAction(self.driver).press(x=1000, y=1530).release().perform()

    def set_form(self, company_info, count, q):
        data = company_info
        q.put(data['company_name'])
        if count != 1:
            self.re_in_form()
        time.sleep(2)
        # 楼宇名称
        building_name = self.find_element_util(
            "//android.view.View[@text='01*' or @text='*01']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        building_name.clear()
        building_name.send_keys(data['building_name'])
        self.swip_up(1)
        # 楼宇ID
        building_id = self.find_element_util(
            "//android.view.View[@text='02*' or @text='*02']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        building_id.clear()
        building_id.send_keys(data['building_id'])
        self.swip_up(1)
        # 客户单位名称
        company_name = self.find_element_util(
            "//android.view.View[@text='03*' or @text='*03']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        company_name.clear()
        time.sleep(0.2)
        company_name.send_keys(data['company_name'])
        self.swip_up(1)
        # self.swip_up(1)
        # 客户联系人
        contacts_name = self.find_element_util(
            "//android.view.View[@text='04*' or @text='*04']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        contacts_name.clear()
        time.sleep(0.2)
        contacts_name.send_keys(data['customer_company_name'])
        self.swip_up(1)
        # 客户联系电话
        contacts_phone = self.find_element_util(
            "//android.view.View[@text='05*' or @text='*05']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        contacts_phone.clear()
        time.sleep(0.2)
        contacts_phone.send_keys(data['contacts_phone'])
        self.swip_up(1)
        # 客户类型
        contacts_type = self.find_element_util(
            "//android.widget.RadioButton[@text='潜在客户']")
        contacts_type.click()
        self.swip_up(1)
        # 使用竞争对手业务套餐名称
        competitor_package = self.find_element_util(
            "//android.view.View[@text='07*' or @text='*07']/following-sibling::android.view.View[1]/android.view.View[2]/android.widget.EditText")
        competitor_package.clear()
        time.sleep(0.2)
        competitor_package.send_keys(data['competitor_package'])
        self.swip_up(1)
        # 使用竞争对手业务资费
        competitor_money = self.find_element_util(
            "//android.view.View[@text='08*' or @text='*08']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        competitor_money.clear()
        time.sleep(0.2)
        competitor_money.send_keys(data['competitor_money'])
        self.swip_up(1)
        # 客户套餐到期日期
        competitor_end_time = self.find_element_util(
            "//android.view.View[@text='09*' or @text='*09']/following-sibling::android.view.View[1]/android.view.View[2]/android.view.View")
        competitor_end_time.click()
        # 在日历上选择时间
        self.select_time2()
        self.swip_up(1)
        # 推荐业务
        recommend_info = self.find_element_util(
            "//android.view.View[@text='10*' or @text='*10']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        recommend_info.send_keys(data['recommend_info'])
        self.swip_up(1)
        # 预计月收入
        estimated_income = self.find_element_util(
            "//android.view.View[@text='11*' or @text='*11']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        estimated_income.clear()
        time.sleep(0.2)
        estimated_income.send_keys(data['estimated_income'])
        self.swip_up(1)
        # 预采取营销措施
        marketing = self.find_element_util(
            "//android.view.View[@text='12*' or @text='*12']/following-sibling::android.view.View[1]/android.view.View[2]/android.widget.EditText")
        marketing.clear()
        time.sleep(0.2)
        marketing.send_keys(data['marketing'])
        self.swip_up(1)
        # 拜访总结
        summary = self.find_element_util(
            "//android.view.View[@text='13*' or @text='*13']/following-sibling::android.view.View[1]/android.view.View[3]/android.widget.EditText")
        summary.clear()
        time.sleep(0.4)
        summary.send_keys(data['summary'])
        time.sleep(0.5)
        self.swip_up(1)
        # 位置
        location = self.find_element_util(
            "//android.view.View[@text='15*' or @text='*15']/following-sibling::android.view.View[1]/android.widget.Button")
        location.click()
        # 位置确认
        location_ack = self.find_element_util("com.tencent.wework:id/kv7", "ID")
        while not location_ack.is_enabled():
            time.sleep(0.5)
        location_ack.click()
        self.swip_up(1)
        # 点击确认android.widget.Button
        submit_button = self.find_element_util(
            "//android.view.View[@text='15*' or @text='*15']/following-sibling::android.widget.Button[1]")
        # 提交数据
        submit_button.click()
        time.sleep(0.5)
        # 点击确认按钮
        ack_button = self.find_element_util(
            "//android.widget.Button[@text='确认']"
        )
        ack_button.click()
        self.save_img("ok")
        time.sleep(int(self.reader.get("run_data", "time_sleep")))
        self.driver.back()
        self.driver.back()

    # 运行appium脚本主程序
    def run_main(self, q, arr_data_list=None):
        if arr_data_list is None:
            # 读取表格信息
            company_list = read_excl(1077, 30)
            # 整理表格数据
            arr_data_list = arrangement_data(company_list)
        # 查找到填写信息的表格
        self.find_table()
        start_number = self.reader.get("run_data", "start_num")
        count = 1
        # 遍历上传数据
        for data in arr_data_list:
            q.put('执行第 %d 条数据' % count)
            start_time = int(time.time())
            self.set_form(data, count, q)
            use_time = int(time.time() - start_time)
            logger.info(data)
            row = str(int(start_number) + count - 1)
            # 写日志
            logger.info("执行%d次，第 %s 行数据，耗时 %d 秒" % (count, row, use_time))
            # 加入队列用作显示
            q.put('执行%d次，第 %s 行数据，耗时 %d 秒' % (count, row, use_time))
            # 修改配置文件中数据
            self.writer.write("run_data", "start_num", str(int(start_number) + count))
            count += 1
        q.put('执行结束…感谢您的使用！再见')

    # 运行appium脚本主程序 测试程序
    def run_main_test(self):
        # 读取表格信息
        company_list = read_excl(1143, 2)
        # 整理表格数据
        arr_data_list = arrangement_data(company_list, "COSMO财富中心一栋", "5014805079")
        # 查找到填写信息的表格
        self.find_table()
        count = 1
        # 遍历上传数据
        for data in arr_data_list:
            start_time = int(time.time())
            # self.set_form(data, count)
            use_time = int(time.time() - start_time)
            logger.info(data)
            logger.info("执行第 %d 次完成，耗时 %d" % (count, use_time))
            count += 1


def test_shuju():
    SetInfo().run_main1()


if __name__ == '__main__':
    # SetInfo().run_main()
    # # self.driver.send_sms()
    # # self.driver.quit()
    # test_path()
    test_shuju()
