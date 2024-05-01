from appium import webdriver
import subprocess
import time

class AppiumSetup:
    def __init__(self):
        self.driver = None

    def start_appium_session(self):
        # Khởi động ứng dụng bằng cách sử dụng subprocess để chạy ứng dụng từ dòng lệnh.
        subprocess.run(['start', '', 'C:\\Users\\Thanh Truc\\Desktop\\Infomed HIS-LK WAN.appref-ms'], shell=True)

        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(5)

        desired_caps = {
            "app": "Root",
            "platformName": "Windows",
            "deviceName": "WindowsPC"
        }

        self.driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities=desired_caps
        )

    def get_driver(self):
        return self.driver

    def find_element_by_accessibility_id(self, accessibility_id):
        return self.driver.find_element_by_accessibility_id(accessibility_id)

    def find_element_by_name(self, name):
        return self.driver.find_element_by_name(name)

    def find_element_by_xpath(self, xpath):
        return self.driver.find_element_by_xpath(xpath)

    def quit_appium_session(self):
        if self.driver:
            self.driver.quit()
