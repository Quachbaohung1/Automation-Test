import datetime
import math
import pyautogui
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from appium.webdriver.common.mobileby import MobileBy
from Driver.init_driver import AppiumSetup
# from Datatest.field_login import account, password, visit_department, visit_room
import time
import pandas as pd

class function_examination():
    def __init__(self):
        self.appium_setup = AppiumSetup()
        self.wait = WebDriverWait(self.appium_setup, 10)
    #Hàm tạo thông tin bệnh nhân
    def patient_information(self, patient_code):
        # Hiển thị thông tin bệnh nhân
        search_content = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtSearchContent")))
        search_content.click()
        search_content.send_keys(str(int(patient_code)))
        pyautogui.press('enter')
        time.sleep(5)
    def patient_import(self, patient_sysmtom, patient_icd_preliminary, patient_icd_main, patient_icd_attachment, patient_note):
        # Nhập triệu chứng
        patient_sysmtom_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtSymptom")))
        patient_sysmtom_field.click()
        patient_sysmtom_field.send_keys(patient_sysmtom)
        # Nhập ICD sơ bộ
        patient_icd_preliminary_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboInitialDxICD")))
        patient_icd_preliminary_field.click()
        search_icd_preliminary = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_icd_preliminary.click()
        search_icd_preliminary.send_keys(patient_icd_preliminary)
        chooser_search_icd_preliminary = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Mã row0")))
        chooser_search_icd_preliminary.click()
        # Nhập ICD chính
        patient_icd_main_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboDxIcd")))
        patient_icd_main_field.click()
        search_icd_main = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_icd_main.click()
        search_icd_main.send_keys(patient_icd_main)
        chooser_search_icd_main = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Mã row0")))
        chooser_search_icd_main.click()
        # Nhập ICD kèm theo
        patient_icd_attachment_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Mã ICD kèm theo new item row")))
        patient_icd_attachment_field.click()
        search_icd_attachment = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_icd_attachment.click()
        search_icd_attachment.send_keys(patient_icd_attachment)
        # time.sleep(3)
        chooser_search_icd_attachment = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Mã row0")))
        chooser_search_icd_attachment.click()
        #Nhập ghi chú
        patient_note_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtDxNotes")))
        patient_note_field.click()
        patient_note_field.send_keys(patient_note)
        # Click btn Lưu
        button_save = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnSave")))
        button_save.click()
        time.sleep(5)
    #Hàm kê đơn thuốc
    def prescription(self, patient_warehouse, patient_medical, morning, noon, afternoon, evening, patient_advice, patient_note):
        #Mở màn hình kê toa(F3)
        prescription_screen = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Toa thuốc (F3)")))
        prescription_screen.click()
        time.sleep(5)
        #Chọn kho
        warehouse_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboChooseStore")))
        warehouse_field.click()
        if patient_warehouse == "Nhà thuốc 1":
            chooser_warehouse = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
            chooser_warehouse.click()
        if patient_warehouse == "Nhà thuốc - Dịch vụ":
            chooser_warehouse = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row1")))
            chooser_warehouse.click()
        # Nhập ngày dùng thuốc
        medication_day_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtPxDays")))
        medication_day_field.click()
        medication_day_field.send_keys("7")
        # Chọn thuốc
        patient_medical_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Thuốc row0")))
        patient_medical_field.click()
        search_patient_medical = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_patient_medical.click()
        search_patient_medical.send_keys(patient_medical)
        chooser_search_icd_attachment = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Mã row0")))
        chooser_search_icd_attachment.click()
        # Nhập sáng/trưa/chiều/tối
        morning_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Sáng row0")))
        morning_field.click()
        morning_field.send_keys(morning)
        noon_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Trưa row0")))
        noon_field.click()
        noon_field.send_keys(noon)
        afternoon_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Chiều row0")))
        afternoon_field.click()
        afternoon_field.send_keys(afternoon)
        evening_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tối row0")))
        evening_field.click()
        evening_field.send_keys(evening)
        # Nhập cách dùng
        use_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Cách dùng row0")))
        use_field.click()
        use_field.send_keys("usa")
        chooser_use = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Ký hiệu row0")))
        chooser_use.click()
        # Nhập lời dặn
        patient_advice_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtAdvice")))
        patient_advice_field.click()
        patient_advice_field.send_keys(patient_advice)
        # Nhập ghi chú
        patient_notes_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtNotes")))
        patient_notes_field.click()
        patient_notes_field.send_keys(patient_note)
        # Click btn Lưu
        button_save = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnSave")))
        button_save.click()
        time.sleep(5)
    # Hàm chỉ định cận lâm sàng yêu cầu chi tiết
    def perform_paraclinical_single(self, patient_med_servicecode):
        # Mở màn hình chỉ định cận lâm sàng
        perform_paraclinical_screen = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnAddMedService")))
        perform_paraclinical_screen.click()
        time.sleep(5)
        # Nhập mã dịch vụ
        patient_med_servicecode_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtMedServiceCode")))
        patient_med_servicecode_field.click()
        patient_med_servicecode_field.send_keys(patient_med_servicecode)
        chooser_price = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtPrice")))
        chooser_price.click()
        # Thêm dịch vụ vào lưới
        button_addF2 = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnAddToGrid")))
        button_addF2.click()
        # Xóa dịch vụ
        button_cancle = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnCancel")))
        button_cancle.click()
        # Lưu chỉ định dịch vụ
        button_save = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnSave")))
        button_save.click()
        # Tắt màn hình chỉ định dịch vụ
        button_close = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnClose")))
        button_close.click()
        time.sleep(5)
    # Hàm chỉ định cận lâm sàng yêu cầu danh sách
    def perform_paraclinical_list(self, patient_med_servicename):
        # Mở màn hình chỉ định cận lâm sàng
        perform_paraclinical_screen = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnAddMedService")))
        perform_paraclinical_screen.click()
        time.sleep(5)
        # Chọn tab Yêu cầu danh sách
        tab_list = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Yêu cầu danh sách")))
        tab_list.click()
        # Nhập mã dịch vụ
        patient_med_servicename_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên dịch vụ filter row")))
        patient_med_servicename_field.click()
        patient_med_servicename_field.send_keys(patient_med_servicename)
        # Thêm dịch vụ vào lưới
        chooser_service = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Thêm row0")))
        chooser_service.click()
        # Lưu chỉ định dịch vụ
        button_save = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnSave")))
        button_save.click()
        # Tắt màn hình chỉ định dịch vụ
        button_close = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnClose")))
        button_close.click()
        time.sleep(5)
    # Hàm kiểm tra các dịch vụ đã được chỉ định
    def perform_paraclinical_check(self):
        # Mở màn hình chỉ định cận lâm sàng
        perform_paraclinical_screen = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnAddMedService")))
        perform_paraclinical_screen.click()
        time.sleep(5)
        # Chọn tab Yêu cầu danh sách
        tab_list = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Danh sách phiếu chỉ định (F6)")))
        tab_list.click()
        # Tắt màn hình chỉ định dịch vụ
        button_close = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnClose")))
        button_close.click()
        time.sleep(5)
    # Hàm phiếu nhập viện
    def hospitalize(self, patient_reason, patient_specialist, patient_department, patient_note, patient_pathological_process, patient_personal_history, patient_family_history, patient_general_examination, patient_specialist_examination, patient_clinial_result):
        # Mở màn hình Phiếu vào viện (F9)
        prescription_screen = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Phiếu vào viện (F9)")))
        prescription_screen.click()
        time.sleep(5)
        # Nhập lý do vào viện
        patient_reason_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtAdmReason")))
        patient_reason_field.click()
        patient_reason_field.send_keys(patient_reason)
        # Nhập khoa và chọn khoa
        chooser_specialist = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboMedRcdType")))
        chooser_specialist.click()
        if patient_specialist == "Nội trú":
            patient_specialist_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 1")))
            patient_specialist_field.click()
        else:
            patient_specialist_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 0")))
            patient_specialist_field.click()
        chooser_department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboAdmWardId")))
        chooser_department.click()
        search_patient_department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_patient_department.click()
        search_patient_department.send_keys(patient_department)
        chooser_patient_department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Name row0")))
        chooser_patient_department.click()
        # Nhập ghi chú
        patient_note_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtWardAdmCertRemark")))
        patient_note_field.click()
        patient_note_field.send_keys(patient_note)
        # Nhập quá trình bệnh lý
        patient_pathological_process_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtPaProcess")))
        patient_pathological_process_field.click()
        patient_pathological_process_field.send_keys(patient_pathological_process)
        # Nhập tiền sử bản thân
        patient_personal_history_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtMedHistPersonClone")))
        patient_personal_history_field.click()
        patient_personal_history_field.send_keys(patient_personal_history)
        # Nhập tiền sử gia đình
        patient_family_history_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtMedHistFamilyClone")))
        patient_family_history_field.click()
        patient_family_history_field.send_keys(patient_family_history)
        # Nhập khám toàn thân
        patient_general_examination_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtGeneralExam")))
        patient_general_examination_field.click()
        patient_general_examination_field.send_keys(patient_general_examination)
        # Nhập khám các bộ phận
        patient_specialist_examination_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtSpecialistExam")))
        patient_specialist_examination_field.click()
        patient_specialist_examination_field.send_keys(patient_specialist_examination)
        # Nhập tóm tắt kết quả cận lâm sàng
        patient_clinial_result_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtSubClinialResult")))
        patient_clinial_result_field.click()
        patient_clinial_result_field.send_keys(patient_clinial_result)
        # Nhập đã xử lý
        button_cls = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnCLS")))
        button_cls.click()
        time.sleep(5)
        # Click btn Lưu
        button_save = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnSave")))
        button_save.click()
    def test_examination_checkin(self):
        # Click chọn Khoa
        department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboWards")))
        department.click()
        search_department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_department.click()
        search_department.send_keys("Khoa Nội tổng hợp")
        chooser_search_department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "gridColumn1 row0")))
        chooser_search_department.click()

        # Click chọn Phòng
        room = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboWardUnits")))
        room.click()
        search_room = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_room.click()
        search_room.send_keys("PK Nội Tiết 1")
        chooser_search_room = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "gridColumn2 row0")))
        chooser_search_room.click()
        # Click btn Xác nhận
        confirmed = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnLogin")))
        confirmed.click()
        # Click chức năng Khám bệnh
        examination_click = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnExamination")))
        examination_click.click()
        # Mở màn hình khám bệnh
        examination_menu = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Khám bệnh")))
        time.sleep(5)
        examination_menu.click()
        examination_item = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.XPATH, "//Button[@Name = 'Khám bệnh']")))
        examination_item.click()
    def test_examination_import(self, TC_ID):
        # Đọc file data test
        file_path = 'D:/HIS automation/Datatest/datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lọc dữ liệu cho TC_ID được truyền vào
        df = df[df['TC_ID'] == TC_ID]
        # Lặp qua từng hàng trong DataFrame và nhập dữ liệu vào trường văn bản trên giao diện người dùng
        for index, row in df.iterrows():
            patient_code = row['patient_code']
            patient_sysmtom = row['patient_sysmtom']
            patient_icd_preliminary = row['patient_icd_preliminary']
            patient_icd_main = row['patient_icd_main']
            patient_icd_attachment = row['patient_icd_attachment']
            patient_note = row['patient_note']

            # Tiếp tục làm việc với ứng dụng bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
            element_visit = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "pnMainPage")))
            element_visit.click()
            # Thông tin bệnh nhân
            function_examination.patient_information(self, patient_code)
            function_examination.patient_import(self, patient_sysmtom, patient_icd_preliminary, patient_icd_main, patient_icd_attachment, patient_note)

    def test_examination_perform_paraclinical_single(self, TC_ID):
        # Đọc file data test
        file_path = 'D:/HIS automation/Datatest/datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lọc dữ liệu cho TC_ID được truyền vào
        df = df[df['TC_ID'] == TC_ID]
        # Lặp qua từng hàng trong DataFrame và nhập dữ liệu vào trường văn bản trên giao diện người dùng
        for index, row in df.iterrows():
            patient_code = row['patient_code']
            patient_sysmtom = row['patient_sysmtom']
            patient_icd_preliminary = row['patient_icd_preliminary']
            patient_icd_main = row['patient_icd_main']
            patient_icd_attachment = row['patient_icd_attachment']
            patient_note = row['patient_note']
            patient_med_servicecode = row['patient_med_servicecode']
        # Thông tin bệnh nhân
        function_examination.patient_information(self, patient_code)
        function_examination.patient_import(self, patient_sysmtom, patient_icd_preliminary, patient_icd_main,
                                                patient_icd_attachment, patient_note)
        # Chỉ định cận lâm sàng tab yêu cầu chi tiết
        function_examination.perform_paraclinical_single(self, patient_med_servicecode)
        # Kiểm tra các dịch vụ cận lâm sàng đã chỉ định
        function_examination.perform_paraclinical_check(self)
    def test_examination_perform_paraclinical_list(self, TC_ID):
        # Đọc file data test
        file_path = 'D:/HIS automation/Datatest/datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lọc dữ liệu cho TC_ID được truyền vào
        df = df[df['TC_ID'] == TC_ID]
        # Lặp qua từng hàng trong DataFrame và nhập dữ liệu vào trường văn bản trên giao diện người dùng
        for index, row in df.iterrows():
            patient_code = row['patient_code']
            patient_sysmtom = row['patient_sysmtom']
            patient_icd_preliminary = row['patient_icd_preliminary']
            patient_icd_main = row['patient_icd_main']
            patient_icd_attachment = row['patient_icd_attachment']
            patient_note = row['patient_note']
            patient_med_servicename = row['patient_med_servicename']
        # Thông tin bệnh nhân
        function_examination.patient_information(self, patient_code)
        function_examination.patient_import(self, patient_sysmtom, patient_icd_preliminary, patient_icd_main,
                                                patient_icd_attachment, patient_note)
        # Chỉ định cận lâm sàng tab yêu cầu danh sách
        function_examination.perform_paraclinical_list(self, patient_med_servicename)
        # Kiểm tra các dịch vụ cận lâm sàng đã chỉ định
        function_examination.perform_paraclinical_check(self)
    def test_prescription(self, TC_ID):
        # Đọc file data test
        file_path = 'D:/HIS automation/Datatest/datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lọc dữ liệu cho TC_ID được truyền vào
        df = df[df['TC_ID'] == TC_ID]
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
        # Thông tin bệnh nhân
        function_examination.patient_information(self, patient_code)
        function_examination.patient_import(self, patient_sysmtom, patient_icd_preliminary, patient_icd_main,
                                            patient_icd_attachment, patient_note)
        # Kê toa thuốc
        function_examination.prescription(self, patient_warehouse, patient_medical, morning, noon, afternoon, evening,
                                          patient_advice, patient_note)
    def test_hospitalize(self, TC_ID):
        # Đọc file data test
        file_path = 'D:/HIS automation/Datatest/datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lọc dữ liệu cho TC_ID được truyền vào
        df = df[df['TC_ID'] == TC_ID]
        # Lặp qua từng hàng trong DataFrame và nhập dữ liệu vào trường văn bản trên giao diện người dùng
        for index, row in df.iterrows():
            patient_code = row['patient_code']
            patient_sysmtom = row['patient_sysmtom']
            patient_icd_preliminary = row['patient_icd_preliminary']
            patient_icd_main = row['patient_icd_main']
            patient_icd_attachment = row['patient_icd_attachment']
            patient_note = row['patient_note']
            patient_reason = row['patient_reason']
            patient_specialist = row['patient_specialist']
            patient_department = row['patient_department']
            patient_pathological_process = row['patient_pathological_process']
            patient_personal_history = row['patient_personal_history']
            patient_family_history = row['patient_family_history']
            patient_general_examination = row['patient_general_examination']
            patient_specialist_examination = row['patient_specialist_examination']
            patient_clinial_result = row['patient_clinial_result']

        # Thông tin bệnh nhân
        function_examination.patient_information(self, patient_code)
        function_examination.patient_import(self, patient_sysmtom, patient_icd_preliminary, patient_icd_main,
                                            patient_icd_attachment, patient_note)
        # Chỉ định nhập viện cho bệnh nhân
        function_examination.hospitalize(self, patient_reason, patient_specialist, patient_department, patient_note,
                                         patient_pathological_process, patient_personal_history, patient_family_history,
                                         patient_general_examination, patient_specialist_examination,
                                         patient_clinial_result)