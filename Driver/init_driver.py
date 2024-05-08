from appium import webdriver
import subprocess
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

    def find_element(self, by, value):
        return self.driver.find_element(by, value)

    def wait_for_element(self, by, value, timeout=10):
        return WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )

    def quit_appium_session(self):
        if self.driver:
            self.driver.quit()
