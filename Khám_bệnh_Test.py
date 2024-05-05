import unittest
import pandas as pd
import pyautogui
import time
from Tiếp_nhận_Test import function_visit
from init_driver import AppiumSetup

class function_examination():
    def __init__(self):
        self.appium_setup = AppiumSetup()
    def patient_information(self, patient_code, patient_sysmtom):
        # Hiển thị thông tin bệnh nhân
        search_content = self.appium_setup.find_element_by_accessibility_id("txtSearchContent")
        search_content.click()
        search_content.send_keys(patient_code)
        pyautogui.press('enter')
        time.sleep(5)

        # Nhập triệu chứng
        patient_sysmtom_field = self.appium_setup.find_element_by_accessibility_id("txtSymptom")
        patient_sysmtom_field.click()
        patient_sysmtom_field.send_keys(patient_sysmtom)

        # Nhập ICD sơ bộ

        # Nhập ICD chính

        # Nhập chẩn đoán

        # Nhập ICD kèm theo

        # Click btn Lưu

    def prescription(self):

    def perform_paraclinical(self):

    def hospitalize(self):



class ExaminationTestSetup(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.appium_setup = AppiumSetup()
        cls.appium_setup.start_appium_session()

    @classmethod
    def tearDownClass(cls):
        cls.appium_setup.quit_appium_session()

    def test_examination_checkin(self):
        # Tạo một instance
        function_visit.login_with_parameters(self)
        # Click chọn Khoa
        department = self.appium_setup.find_element_by_accessibility_id("cboWards")
        self.assertIsNotNone(department, "Không tìm thấy phần tử bằng accessibility_id")
        department.click()
        search_department = self.appium_setup.find_element_by_accessibility_id("teFind")
        self.assertIsNotNone(search_department, "Không tìm thấy phần tử bằng accessibility_id")
        search_department.click()
        search_department.send_keys("Khoa Nội tổng hợp")
        chooser_search_department = self.appium_setup.find_element_by_name("gridColumn1 row0")
        self.assertIsNotNone(chooser_search_department, "Không tìm thấy phần tử bằng name")
        chooser_search_department.click()

        # Click chọn Phòng
        room = self.appium_setup.find_element_by_accessibility_id("cboWardUnits")
        self.assertIsNotNone(room, "Không tìm thấy phần tử bằng accessibility_id")
        room.click()
        search_room = self.appium_setup.find_element_by_accessibility_id("teFind")
        self.assertIsNotNone(search_room, "Không tìm thấy phần tử bằng accessibility_id")
        search_room.click()
        search_room.send_keys("PK Nội Tiết 1")
        chooser_search_room = self.appium_setup.find_element_by_name("gridColumn2 row0")
        self.assertIsNotNone(chooser_search_room, "Không tìm thấy phần tử bằng name")
        chooser_search_room.click()

        # Click btn Xác nhận
        confirmed = self.appium_setup.find_element_by_accessibility_id("btnLogin")
        self.assertIsNotNone(confirmed, "Không tìm thấy phần tử bằng accessibility_id")
        confirmed.click()

        # Click chức năng Khám bệnh
        examination_click = self.appium_setup.find_element_by_accessibility_id("btnExamination")
        self.assertIsNotNone(examination_click, "Không tìm thấy phần tử bằng accessibility_id")
        examination_click.click()
        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(5)

        # Mở màn hình khám bệnh
        examination_menu = self.appium_setup.find_element_by_name("Khám bệnh")
        self.assertIsNotNone(examination_menu, "Không tìm thấy phần tử bằng name")
        examination_menu.click()
        examination_item = self.appium_setup.find_element_by_xpath("//Button[@Name = 'Khám bệnh' ]")
        self.assertIsNotNone(examination_item, "Không tìm thấy phần tử bằng xpath")
        examination_item.click()
        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(5)

    def test_examination_function(self):
        # Đọc file data test
        file_path = 'datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lặp qua từng hàng trong DataFrame và nhập dữ liệu vào trường văn bản trên giao diện người dùng
        for index, row in df.iterrows():
            patient_code = row['patient_code']
            patient_sysmtom = row['patient_sysmtom']

            # Tiếp tục làm việc với ứng dụng bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
            element_visit = self.appium_setup.find_element_by_accessibility_id("pnMainPage")
            element_visit.click()
            self.assertIsNotNone(element_visit, "Không tìm thấy phần tử bằng accessibility_id")

            #Thông tin bệnh nhân
            function_examination.patient_information(self, patient_code , patient_sysmtom)

            #Chỉ định cận lâm sàng

            #Kê toa thuốc





if __name__ == "__main__":
    unittest.main()
