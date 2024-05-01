import unittest
from init_driver import AppiumSetup

class Automation_Test(unittest.TestCase):
    def __init__(self):
        self.appium_setup = AppiumSetup()

    def start_application(self):
        # Start Appium session
        self.appium_setup.start_appium_session()

    def get_driver(self):
        return self.appium_setup.get_driver()

    def quit_application(self):
        # Quit Appium session
        self.appium_setup.quit_appium_session()

    def find_element_by_accessibility_id(self, accessibility_id):
        return self.appium_setup.find_element_by_accessibility_id(accessibility_id)

    def find_element_by_name(self, name):
        return self.appium_setup.find_element_by_name(name)

    def setUp(self):
        self.automation_test = Automation_Test()
        self.automation_test.start_application()

    def tearDown(self):
        self.automation_test.quit_application()

    def main(self):
        self.start_application()
        # Tiếp tục làm việc với ứng dụng của bạn bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
        element = self.find_element_by_accessibility_id("FrmMain")
        element.click()

        #Nhập Tài khoản
        account_field = self.find_element_by_accessibility_id("txtAccName")
        account_field.click()  # Click vào trường "Tài khoản" để đảm bảo trường đóng được chọn
        # Xóa nội dung trường "Tài khoản" trước khi nhập liệu mới
        account_field.clear()
        # Nhập dữ liệu mới vào trường "Tài khoản"
        account_field.send_keys("hungqb")

        #Nhập Mật khẩu
        password_field = self.find_element_by_accessibility_id("txtPassword")
        password_field.click()  # Click vào trường "Tài khoản" để đảm bảo trường đóng được chọn
        # Xóa nội dung trường "Tài khoản" trước khi nhập liệu mới
        password_field.clear()
        # Nhập dữ liệu mới vào trường "Tài khoản"
        password_field.send_keys("1")

        #Click btn Đăng nhập
        login = self.find_element_by_accessibility_id("btnLogin")
        login.click()

        #Click chọn Khoa
        department = self.find_element_by_accessibility_id("cboWards")
        department.click()
        search_department = self.find_element_by_accessibility_id("teFind")
        search_department.click()
        search_department.send_keys("Phòng điều dưỡng")
        chooser_search_department = self.find_element_by_name("gridColumn1 row0")
        chooser_search_department.click()

        #Click chọn Phòng
        room = self.find_element_by_accessibility_id("cboWardUnits")
        room.click()
        search_room = self.find_element_by_accessibility_id("teFind")
        search_room.click()
        search_room.send_keys("Tiếp nhận khu công")
        chooser_search_room = self.find_element_by_name("gridColumn2 row0")
        chooser_search_room.click()

        # Click btn Xác nhận
        confirmed = self.find_element_by_accessibility_id("btnLogin")
        confirmed.click()

if __name__ == "__main__":
    your_class_instance = Automation_Test()
    your_class_instance.main()
    unittest.main()
