import unittest
import pandas as pd
import pyautogui
import time
from Tiếp_nhận_Test import function_visit
from init_driver import AppiumSetup

class function_examination():
    def __init__(self):
        self.appium_setup = AppiumSetup()
    #Hàm tạo thông tin bệnh nhân
    def patient_information(self, patient_code, patient_sysmtom, patient_icd_preliminary, patient_icd_main, patient_icd_attachment, patient_note):
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
        patient_icd_preliminary_field = self.appium_setup.find_element_by_accessibility_id("cboInitialDxICD")
        patient_icd_preliminary_field.click()
        search_icd_preliminary = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_icd_preliminary.click()
        search_icd_preliminary.send_keys(patient_icd_preliminary)
        chooser_search_icd_preliminary = self.appium_setup.find_element_by_name("Mã row0")
        chooser_search_icd_preliminary.click()
        # Nhập ICD chính
        patient_icd_main_field = self.appium_setup.find_element_by_accessibility_id("cboDxIcd")
        patient_icd_main_field.click()
        search_icd_main = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_icd_main.click()
        search_icd_main.send_keys(patient_icd_main)
        chooser_search_icd_main = self.appium_setup.find_element_by_name("Mã row0")
        chooser_search_icd_main.click()
        # Nhập ICD kèm theo
        patient_icd_attachment_field = self.appium_setup.find_element_by_name("Mã ICD kèm theo new item row")
        patient_icd_attachment_field.click()
        search_icd_attachment = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_icd_attachment.click()
        search_icd_attachment.send_keys(patient_icd_attachment)
        chooser_search_icd_attachment = self.appium_setup.find_element_by_name("Mã row0")
        chooser_search_icd_attachment.click()
        #Nhập ghi chú
        patient_note_field = self.appium_setup.find_element_by_accessibility_id("txtDxNotes")
        patient_note_field.click()
        patient_note_field.send_keys(patient_note)
        # Click btn Lưu
        button_save = self.appium_setup.find_element_by_accessibility_id("btnSave")
        button_save.click()
    #Hàm kê đơn thuốc
    def prescription(self, patient_warehouse, patient_medical, morning, noon, afternoon, evening, patient_advice, patient_note):
        #Mở màn hình kê toa(F3)
        prescription_screen = self.appium_setup.find_element_by_name("Toa thuốc (F3)")
        prescription_screen.click()
        time.sleep(5)
        #Chọn kho
        warehouse_field = self.appium_setup.find_element_by_accessibility_id("cboChooseStore")
        warehouse_field.click()
        if patient_warehouse == "Nhà thuốc 1":
            chooser_warehouse = self.appium_setup.find_element_by_name("Tên row0")
            chooser_warehouse.click()
        if patient_warehouse == "Nhà thuốc - Dịch vụ":
            chooser_warehouse = self.appium_setup.find_element_by_name("Tên row1")
            chooser_warehouse.click()
        #Nhập ngày dùng thuốc
        medication_day_field = self.appium_setup.find_element_by_accessibility_id("txtPxDays")
        medication_day_field.click()
        medication_day_field.send_keys("7")
        #Chọn thuốc
        patient_medical_field = self.appium_setup.find_element_by_name("Thuốc row0")
        patient_medical_field.click()
        search_patient_medical = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_patient_medical.click()
        search_patient_medical.send_keys(patient_medical)
        chooser_search_icd_attachment = self.appium_setup.find_element_by_name("Mã row0")
        chooser_search_icd_attachment.click()
        #Nhập sáng/trưa/chiều/tối
        morning_field = self.appium_setup.find_element_by_name("Sáng row0")
        morning_field.click()
        morning_field.send_keys(morning)
        noon_field = self.appium_setup.find_element_by_name("Trưa row0")
        noon_field.click()
        noon_field.send_keys(noon)
        afternoon_field = self.appium_setup.find_element_by_name("Chiều row0")
        afternoon_field.click()
        afternoon_field.send_keys(afternoon)
        evening_field = self.appium_setup.find_element_by_name("Tối row0")
        evening_field.click()
        evening_field.send_keys(evening)
        #Nhập cách dùng
        use_field = self.appium_setup.find_element_by_name("Cách dùng row0")
        use_field.click()
        use_field.send_keys("usa")
        chooser_use = self.appium_setup.find_element_by_name("Ký hiệu row0")
        chooser_use.click()
        #Nhập lời dặn
        patient_advice_field = self.appium_setup.find_element_by_accessibility_id("txtAdvice")
        patient_advice_field.click()
        patient_advice_field.send_keys(patient_advice)
        #Nhập ghi chú
        patient_notes_field = self.appium_setup.find_element_by_accessibility_id("txtNotes")
        patient_notes_field.click()
        patient_notes_field.send_keys(patient_note)
        # Click btn Lưu
        button_save = self.appium_setup.find_element_by_accessibility_id("btnSave")
        button_save.click()
    #Hàm chỉ định cận lâm sàng
    def perform_paraclinical(self, patient_med_servicecode):
        # Mở màn hình kê toa(F3)
        perform_paraclinical_screen = self.appium_setup.find_element_by_name("btnAddMedService")
        perform_paraclinical_screen.click()
        time.sleep(5)
        #Nhập mã dịch vụ
        patient_med_servicecode_field = self.appium_setup.find_element_by_accessibility_id("txtMedServiceCode")
        patient_med_servicecode_field.click()
        patient_med_servicecode_field.send_keys(patient_med_servicecode)
    # #Hàm phiếu nhập viện
    # def hospitalize(self):



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
        department.click()
        search_department = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_department.click()
        search_department.send_keys("Khoa Nội tổng hợp")
        chooser_search_department = self.appium_setup.find_element_by_name("gridColumn1 row0")
        chooser_search_department.click()

        # Click chọn Phòng
        room = self.appium_setup.find_element_by_accessibility_id("cboWardUnits")
        room.click()
        search_room = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_room.click()
        search_room.send_keys("PK Nội Tiết 1")
        chooser_search_room = self.appium_setup.find_element_by_name("gridColumn2 row0")
        chooser_search_room.click()

        # Click btn Xác nhận
        confirmed = self.appium_setup.find_element_by_accessibility_id("btnLogin")
        confirmed.click()

        # Click chức năng Khám bệnh
        examination_click = self.appium_setup.find_element_by_accessibility_id("btnExamination")
        examination_click.click()
        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(10)

        # Mở màn hình khám bệnh
        examination_menu = self.appium_setup.find_element_by_name("Khám bệnh")
        examination_menu.click()
        examination_item = self.appium_setup.find_element_by_xpath("//Button[@Name = 'Khám bệnh' ]")
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
            patient_icd_preliminary = row['patient_icd_preliminary']
            patient_icd_main = row['patient_icd_main']
            patient_icd_attachment = row['patient_icd_attachment']
            patient_note = row['patient_note']
            patient_warehouse = row['patient_warehouse']
            patient_medical = row['patient_medical']
            morning = row['morning']
            noon = row['noon']
            afternoon = row['afternoon']
            evening = row['evening']
            patient_advice = row['patient_advice']
            patient_med_servicecode = row['patient_med_servicecode']


            # Tiếp tục làm việc với ứng dụng bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
            element_visit = self.appium_setup.find_element_by_accessibility_id("pnMainPage")
            element_visit.click()
            #Thông tin bệnh nhân
            function_examination.patient_information(self, patient_code , patient_sysmtom, patient_icd_preliminary, patient_icd_main, patient_icd_attachment, patient_note)
            #Chỉ định cận lâm sàng
            function_examination.perform_paraclinical(self, patient_med_servicecode)
            #Kê toa thuốc
            function_examination.prescription(self, patient_warehouse, patient_medical, morning, noon, afternoon, evening, patient_advice, patient_note)




if __name__ == "__main__":
    unittest.main()
