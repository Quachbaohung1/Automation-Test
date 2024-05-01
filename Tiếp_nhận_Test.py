import unittest
from init_driver import AppiumSetup
import time

class TestAppiumSetup(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.appium_setup = AppiumSetup()
        cls.appium_setup.start_appium_session()

    @classmethod
    def tearDownClass(cls):
        cls.appium_setup.quit_appium_session()

    def test_login(self):
        # Tiếp tục làm việc với ứng dụng của bạn bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
        element = self.appium_setup.find_element_by_accessibility_id("FrmMain")
        element.click()
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")

        # Nhập Tài khoản
        account_field = self.appium_setup.find_element_by_accessibility_id("txtAccName")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        account_field.click()  # Click vào trường "Tài khoản" để đảm bảo trường đóng được chọn
        # Xóa nội dung trường "Tài khoản" trước khi nhập liệu mới
        account_field.clear()
        # Nhập dữ liệu mới vào trường "Tài khoản"
        account_field.send_keys("hungqb")

        # Nhập Mật khẩu
        password_field = self.appium_setup.find_element_by_accessibility_id("txtPassword")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        password_field.click()  # Click vào trường "Tài khoản" để đảm bảo trường đóng được chọn
        # Xóa nội dung trường "Tài khoản" trước khi nhập liệu mới
        password_field.clear()
        # Nhập dữ liệu mới vào trường "Tài khoản"
        password_field.send_keys("1")

        # Click btn Đăng nhập
        login = self.appium_setup.find_element_by_accessibility_id("btnLogin")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        login.click()

        # Click chọn Khoa
        department = self.appium_setup.find_element_by_accessibility_id("cboWards")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        department.click()
        search_department = self.appium_setup.find_element_by_accessibility_id("teFind")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        search_department.click()
        search_department.send_keys("Phòng điều dưỡng")
        chooser_search_department = self.appium_setup.find_element_by_name("gridColumn1 row0")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng name")
        chooser_search_department.click()

        # Click chọn Phòng
        room = self.appium_setup.find_element_by_accessibility_id("cboWardUnits")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        room.click()
        search_room = self.appium_setup.find_element_by_accessibility_id("teFind")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        search_room.click()
        search_room.send_keys("Tiếp nhận khu công")
        chooser_search_room = self.appium_setup.find_element_by_name("gridColumn2 row0")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng name")
        chooser_search_room.click()

        # Click btn Xác nhận
        confirmed = self.appium_setup.find_element_by_accessibility_id("btnLogin")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        confirmed.click()

        #Click chức năng Tiếp nhận
        visit_click = self.appium_setup.find_element_by_accessibility_id("btnVisit")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng accessibility_id")
        visit_click.click()
        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(5)

        #Mở màn hình tiếp nhận
        visit_menu = self.appium_setup.find_element_by_name("Tiếp nhận")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng name")
        visit_menu.click()
        visit_item = self.appium_setup.find_element_by_xpath("//Button[@Name='Tiếp nhận']")
        self.assertIsNotNone(element, "Không tìm thấy phần tử bằng xpath")
        visit_item.click()
        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(5)

    #def test_function(self):



if __name__ == "__main__":
    unittest.main()
