import unittest
import pandas as pd
import pyautogui
import time
from Tiếp_nhận_Test import function_visit
from Driver.init_driver import AppiumSetup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from appium.webdriver.common.mobileby import By

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
        time.sleep(5)
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
        # Nhập ngày dùng thuốc
        medication_day_field = self.appium_setup.find_element_by_accessibility_id("txtPxDays")
        medication_day_field.click()
        medication_day_field.send_keys("7")
        # Chọn thuốc
        patient_medical_field = self.appium_setup.find_element_by_name("Thuốc row0")
        patient_medical_field.click()
        search_patient_medical = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_patient_medical.click()
        search_patient_medical.send_keys(patient_medical)
        chooser_search_icd_attachment = self.appium_setup.find_element_by_name("Mã row0")
        chooser_search_icd_attachment.click()
        # Nhập sáng/trưa/chiều/tối
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
        # Nhập cách dùng
        use_field = self.appium_setup.find_element_by_name("Cách dùng row0")
        use_field.click()
        use_field.send_keys("usa")
        chooser_use = self.appium_setup.find_element_by_name("Ký hiệu row0")
        chooser_use.click()
        # Nhập lời dặn
        patient_advice_field = self.appium_setup.find_element_by_accessibility_id("txtAdvice")
        patient_advice_field.click()
        patient_advice_field.send_keys(patient_advice)
        # Nhập ghi chú
        patient_notes_field = self.appium_setup.find_element_by_accessibility_id("txtNotes")
        patient_notes_field.click()
        patient_notes_field.send_keys(patient_note)
        # Click btn Lưu
        button_save = self.appium_setup.find_element_by_accessibility_id("btnSave")
        button_save.click()
        time.sleep(5)
    # Hàm chỉ định cận lâm sàng yêu cầu chi tiết
    def perform_paraclinical_single(self, patient_med_servicecode):
        # Mở màn hình chỉ định cận lâm sàng
        perform_paraclinical_screen = self.appium_setup.find_element_by_accessibility_id("btnAddMedService")
        perform_paraclinical_screen.click()
        time.sleep(5)
        # Nhập mã dịch vụ
        patient_med_servicecode_field = self.appium_setup.find_element_by_accessibility_id("txtMedServiceCode")
        patient_med_servicecode_field.click()
        patient_med_servicecode_field.send_keys(patient_med_servicecode)
        chooser_price = self.appium_setup.find_element_by_accessibility_id("txtPrice")
        chooser_price.click()
        # Thêm dịch vụ vào lưới
        button_addF2 = self.appium_setup.find_element_by_accessibility_id("btnAddToGrid")
        button_addF2.click()
        # Xóa dịch vụ
        button_cancle = self.appium_setup.find_element_by_accessibility_id("btnCancel")
        button_cancle.click()
        # Lưu chỉ định dịch vụ
        button_save = self.appium_setup.find_element_by_accessibility_id("btnSave")
        button_save.click()
        # Tắt màn hình chỉ định dịch vụ
        button_close = self.appium_setup.find_element_by_accessibility_id("btnClose")
        button_close.click()
        time.sleep(5)
    # Hàm chỉ định cận lâm sàng yêu cầu danh sách
    def perform_paraclinical_list(self, patient_med_servicename):
        # Mở màn hình chỉ định cận lâm sàng
        perform_paraclinical_screen = self.appium_setup.find_element_by_accessibility_id("btnAddMedService")
        perform_paraclinical_screen.click()
        time.sleep(5)
        # Chọn tab Yêu cầu danh sách
        tab_list = self.appium_setup.find_element_by_name("Yêu cầu danh sách")
        tab_list.click()
        # Nhập mã dịch vụ
        patient_med_servicename_field = self.appium_setup.find_element_by_name("Tên dịch vụ filter row")
        patient_med_servicename_field.click()
        patient_med_servicename_field.send_keys(patient_med_servicename)
        # Thêm dịch vụ vào lưới
        chooser_service = self.appium_setup.find_element_by_name("Thêm row0")
        chooser_service.click()
        # Lưu chỉ định dịch vụ
        button_save = self.appium_setup.find_element_by_accessibility_id("btnSave")
        button_save.click()
        # Tắt màn hình chỉ định dịch vụ
        button_close = self.appium_setup.find_element_by_accessibility_id("btnClose")
        button_close.click()
        time.sleep(5)
    # Hàm kiểm tra các dịch vụ đã được chỉ định
    def perform_paraclinical_check(self):
        # Mở màn hình chỉ định cận lâm sàng
        perform_paraclinical_screen = self.appium_setup.find_element_by_accessibility_id("btnAddMedService")
        perform_paraclinical_screen.click()
        time.sleep(5)
        # Chọn tab Yêu cầu danh sách
        tab_list = self.appium_setup.find_element_by_name("Danh sách phiếu chỉ định (F6)")
        tab_list.click()
        # Tắt màn hình chỉ định dịch vụ
        button_close = self.appium_setup.find_element_by_accessibility_id("btnClose")
        button_close.click()
        time.sleep(5)
    # Hàm phiếu nhập viện
    def hospitalize(self, patient_reason, patient_specialist, patient_department, patient_note, patient_pathological_process, patient_personal_history, patient_family_history, patient_general_examination, patient_specialist_examination, patient_clinial_result):
        # Mở màn hình Phiếu vào viện (F9)
        prescription_screen = self.appium_setup.find_element_by_name("Phiếu vào viện (F9)")
        prescription_screen.click()
        time.sleep(5)
        # Nhập lý do vào viện
        patient_reason_field = self.appium_setup.find_element_by_accessibility_id("txtAdmReason")
        patient_reason_field.click()
        patient_reason_field.send_keys(patient_reason)
        # Nhập khoa và chọn khoa
        chooser_specialist = self.appium_setup.find_element_by_accessibility_id("cboMedRcdType")
        chooser_specialist.click()
        if patient_specialist == "Nội trú":
            patient_specialist_field = self.appium_setup.find_element_by_name("Row 1")
            patient_specialist_field.click()
        else:
            patient_specialist_field = self.appium_setup.find_element_by_name("Row 0")
            patient_specialist_field.click()
        chooser_department = self.appium_setup.find_element_by_accessibility_id("cboAdmWardId")
        chooser_department.click()
        search_patient_department = self.appium_setup.find_element_by_accessibility_id("teFind")
        search_patient_department.click()
        search_patient_department.send_keys(patient_department)
        chooser_patient_department = self.appium_setup.find_element_by_name("Name row0")
        chooser_patient_department.click()
        # Nhập ghi chú
        patient_note_field = self.appium_setup.find_element_by_accessibility_id("txtWardAdmCertRemark")
        patient_note_field.click()
        patient_note_field.send_keys(patient_note)
        # Nhập quá trình bệnh lý
        patient_pathological_process_field = self.appium_setup.find_element_by_accessibility_id("txtPaProcess")
        patient_pathological_process_field.click()
        patient_pathological_process_field.send_keys(patient_pathological_process)
        # Nhập tiền sử bản thân
        patient_personal_history_field = self.appium_setup.find_element_by_accessibility_id("txtMedHistPersonClone")
        patient_personal_history_field.click()
        patient_personal_history_field.send_keys(patient_personal_history)
        # Nhập tiền sử gia đình
        patient_family_history_field = self.appium_setup.find_element_by_accessibility_id("txtMedHistFamilyClone")
        patient_family_history_field.click()
        patient_family_history_field.send_keys(patient_family_history)
        # Nhập khám toàn thân
        patient_general_examination_field = self.appium_setup.find_element_by_accessibility_id("txtGeneralExam")
        patient_general_examination_field.click()
        patient_general_examination_field.send_keys(patient_general_examination)
        # Nhập khám các bộ phận
        patient_specialist_examination_field = self.appium_setup.find_element_by_accessibility_id("txtSpecialistExam")
        patient_specialist_examination_field.click()
        patient_specialist_examination_field.send_keys(patient_specialist_examination)
        # Nhập tóm tắt kết quả cận lâm sàng
        patient_clinial_result_field = self.appium_setup.find_element_by_accessibility_id("txtSubClinialResult")
        patient_clinial_result_field.click()
        patient_clinial_result_field.send_keys(patient_clinial_result)
        # Nhập đã xử lý
        button_cls = self.appium_setup.find_element_by_accessibility_id("btnCLS")
        button_cls.click()
        time.sleep(5)
        # Click btn Lưu
        button_save = self.appium_setup.find_element_by_accessibility_id("btnSave")
        button_save.click()

class ExaminationTestSetup(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.appium_setup = AppiumSetup()
        cls.appium_setup.start_appium_session()
        cls.wait = WebDriverWait(cls.appium_setup, 10)

    @classmethod
    def tearDownClass(cls):
        cls.appium_setup.quit_appium_session()

    def test_examination_checkin(self):
        # Tạo một instance
        function_visit.login_information(self)
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
            patient_med_servicename = row['patient_med_servicename']
            patient_reason = row['patient_reason']
            patient_specialist = row['patient_specialist']
            patient_department = row['patient_department']
            patient_pathological_process = row['patient_pathological_process']
            patient_personal_history = row['patient_personal_history']
            patient_family_history = row['patient_family_history']
            patient_general_examination = row['patient_general_examination']
            patient_specialist_examination = row['patient_specialist_examination']
            patient_clinial_result = row['patient_clinial_result']

            # Tiếp tục làm việc với ứng dụng bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
            element_visit = self.appium_setup.find_element_by_accessibility_id("pnMainPage")
            element_visit.click()
            # Thông tin bệnh nhân
            function_examination.patient_information(self, patient_code , patient_sysmtom, patient_icd_preliminary, patient_icd_main, patient_icd_attachment, patient_note)
            # Chỉ định cận lâm sàng tab yêu cầu chi tiết
            function_examination.perform_paraclinical_single(self, patient_med_servicecode)
            # Chỉ định cận lâm sàng tab yêu cầu danh sách
            function_examination.perform_paraclinical_list(self, patient_med_servicename)
            # Kiểm tra các dịch vụ cận lâm sàng đã chỉ định
            function_examination.perform_paraclinical_check(self)
            # Kê toa thuốc
            function_examination.prescription(self, patient_warehouse, patient_medical, morning, noon, afternoon, evening, patient_advice, patient_note)
            # Chỉ định nhập viện cho bệnh nhân
            function_examination.hospitalize(self, patient_reason, patient_specialist, patient_department, patient_note, patient_pathological_process, patient_personal_history, patient_family_history, patient_general_examination, patient_specialist_examination, patient_clinial_result)


if __name__ == "__main__":
    unittest.main()
