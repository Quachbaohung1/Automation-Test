import datetime
import math
import pyautogui
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from appium.webdriver.common.mobileby import MobileBy
from Driver.init_driver import AppiumSetup
from Datatest.field_login import account, password, visit_department, visit_room
import time
import pandas as pd


class function_visit():
    def __init__(self):
        self.appium_setup = AppiumSetup()
        self.wait = WebDriverWait(self.appium_setup, 10)
    def calculate_age(self, patient_bod):
        # Chuyển đổi ngày tháng năm sinh từ định dạng chuỗi sang đối tượng datetime
        patient_bod_datetime = pd.to_datetime(patient_bod, format='%Y-%m-%d')
        today = datetime.datetime.today()
        age = today.year - patient_bod_datetime.year - (
                (today.month, today.day) < (patient_bod_datetime.month, patient_bod_datetime.day))
        return age
    def login_information(self):
        # Tiếp tục làm việc với ứng dụng của bạn bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
        element = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "FrmMain")))
        element.click()
        # Nhập Tài khoản
        account_field = WebDriverWait(self.appium_setup,10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtAccName")))
        account_field.click()
        #account_field.clear()
        account_field.send_keys(account)
        # Nhập Mật khẩu
        password_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtPassword")))
        password_field.click()
        #password_field.clear()
        password_field.send_keys(password)
        # Click btn Đăng nhập
        login = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnLogin")))
        login.click()
    def login_visit(self):
        function_visit.login_information(self)
        # Click chọn Khoa
        department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboWards")))
        department.click()
        search_department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_department.click()
        search_department.send_keys(visit_department)
        chooser_search_department = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "gridColumn1 row0")))
        chooser_search_department.click()
        # Click chọn Phòng
        room = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboWardUnits")))
        room.click()
        search_room = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
        search_room.click()
        search_room.send_keys(visit_room)
        chooser_search_room = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "gridColumn2 row0")))
        chooser_search_room.click()
        # Click btn Xác nhận
        confirmed = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnLogin")))
        confirmed.click()

        #Click chức năng Tiếp nhận
        visit_click = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnVisit")))
        visit_click.click()
        #Mở màn hình tiếp nhận
        visit_menu = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tiếp nhận")))
        time.sleep(5)
        visit_menu.click()
        visit_item = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.XPATH, "//Button[@Name='Tiếp nhận']")))
        visit_item.click()

    def visit_function(self, TC_ID):
        #Đọc file data test
        file_path = 'D:/HIS automation/Datatest/datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lọc dữ liệu cho TC_ID được truyền vào
        df = df[df['TC_ID'] == TC_ID]
        # Lặp qua từng hàng trong DataFrame và nhập dữ liệu vào trường văn bản trên giao diện người dùng
        for index, row in df.iterrows():
            patient_name = row['patient_name']
            patient_gender = row['patient_gender']
            patient_bod = row['patient_bod']
            patient_nationality = row['patient_nationality']
            patient_nation = row['patient_nation']
            patient_job = row['patient_job']
            patient_address = row['patient_address']
            patient_province = row['patient_province']
            patient_ward = row['patient_ward']
            patient_town = row['patient_town']
            patient_benefit = row['patient_benefit']
            patient_bhyt = row['patient_bhyt']
            patient_place = row['patient_place']
            patient_from = row['patient_from']
            patient_to = row['patient_to']
            patient_clinic = row['patient_clinic']
            patient_service = row['patient_service']
            patient_type = row['patient_type']
            patient_company = row['patient_company']
            patient_relative = row['patient_relative']
            patient_relativeName = row['patient_relativeName']
            patient_relativePhone = row['patient_relativePhone']
            patient_relativeAddress = row['patient_relativeAddress']

            # Tiếp tục làm việc với ứng dụng bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
            time.sleep(5)
            element_visit = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "pnMainPage")))
            element_visit.click()
            # Thực hiện các bước để tiếp nhận bệnh nhân với tên và mã bệnh nhân từ file XLSX
            # Nhập họ và tên
            patient_name_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtFullName")))
            patient_name_field.click()
            if patient_name is None or isinstance(patient_name, float) and math.isnan(patient_name):
                # Nếu giá trị là NaN hoặc None, gán giá trị rỗng cho patient_name
                patient_name = ""
            patient_name_field.send_keys(patient_name)
            #Nhập giới tính
            patient_gender_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboGender")))
            patient_gender_field.click()
            patient_gender_field.send_keys(patient_gender)
            #Nhập ngày sinh
            patient_bod_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "dpDOB")))
            patient_bod_field.click()
            # Kiểm tra xem giá trị có phải là NaN hoặc rỗng không
            if pd.isna(patient_bod) or not patient_bod:
                # Nếu giá trị là NaN hoặc rỗng, gửi một khoảng trắng vào trường patient_bod_field
                patient_bod_field.send_keys("")
            else:
                # Nếu giá trị không phải là NaN hoặc rỗng, chuyển đổi ngày sinh sang định dạng mong muốn
                formatted_bod = pd.to_datetime(patient_bod).strftime('%d%m%Y')
                patient_bod_field.send_keys(formatted_bod)
            # Tính toán tuổi của bệnh nhân
            age = function_visit.calculate_age(self, patient_bod)
            # Kiểm tra xem tuổi của bệnh nhân có dưới 6 không
            if age < 6:
                print("Tuổi của bệnh nhân bé hơn 6 tuổi. Cần có người thân.")
                #Nhập mối quan hệ với bệnh nhân
                patient_relative_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboRelativeType")))
                patient_relative_field.click()
                if patient_relative == "Cha":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 1")))
                    chooser_patient_relative.click()
                elif patient_relative == "Mẹ":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 2")))
                    chooser_patient_relative.click()
                elif patient_relative == "Anh em ruột":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 3")))
                    chooser_patient_relative.click()
                elif patient_relative == "Con":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 4")))
                    chooser_patient_relative.click()
                elif patient_relative == "Người bảo hộ":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 5")))
                    chooser_patient_relative.click()
                elif patient_relative == "Vợ-Chồng":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 6")))
                    chooser_patient_relative.click()
                elif patient_relative == "Người đưa đến":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 7")))
                    chooser_patient_relative.click()
                elif patient_relative == "Ông-Bà":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 8")))
                    chooser_patient_relative.click()
                elif patient_relative == "Cháu":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 9")))
                    chooser_patient_relative.click()
                elif patient_relative == "Cô":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 10")))
                    chooser_patient_relative.click()
                elif patient_relative == "Anh":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 11")))
                    chooser_patient_relative.click()
                elif patient_relative == "Em":
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 12")))
                    chooser_patient_relative.click()
                else:
                    chooser_patient_relative = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 13")))
                    chooser_patient_relative.click()
                #Nhập tên người thân
                patient_relativeName_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtRelativeName")))
                patient_relativeName_field.click()
                patient_relativeName_field.send_keys(patient_relativeName)
                #Nhập số điện thoại người thân
                patient_relativePhone_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtRelativePhone")))
                patient_relativePhone_field.click()
                patient_relativePhone_field.send_keys("0" + str(patient_relativePhone))
                #Nhập địa chỉ của người thân
                patient_relativeAddress_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtRelativeAddress")))
                patient_relativeAddress_field.click()
                patient_relativeAddress_field.send_keys(patient_relativeAddress)
            else:
                print("Tuổi của bệnh nhân lớn hơn hoặc bằng 6 tuổi. Không cần có người thân.")
            #Nhập quốc tịch
            patient_nationality_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboNationary")))
            patient_nationality_field.click()
            search_nationality = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
            search_nationality.click()
            if patient_nationality is None or isinstance(patient_nationality, float) and math.isnan(patient_nationality):
                chooser_delete = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btClear")))
                chooser_delete.click()
            else:
                search_nationality.send_keys(patient_nationality)
                chooser_search_nationality = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
                chooser_search_nationality.click()
            #Nếu quốc tịch là VN thì sẽ nhập dân tộc, nếu khác VN thì không cần nhập ở field này
            if patient_nationality == "Việt Nam":
                patient_nation_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboEthnic")))
                patient_nation_field.click()
                patient_nation_field.clear()
                patient_nation_field.send_keys(patient_nation)
            #Nhập nghề nghiệp
            patient_job_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboOccupation")))
            patient_job_field.click()
            patient_job_field.clear()
            patient_job_field.send_keys(patient_job)
            #Nhập địa chỉ
            patient_address_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtAddress")))
            patient_address_field.click()
            if patient_address is None or isinstance(patient_address, float) and math.isnan(patient_address):
                # Nếu giá trị là NaN hoặc None, gán giá trị rỗng cho patient_name
                patient_address = ""
            patient_address_field.send_keys(patient_address)
            #Nhập tỉnh
            patient_province_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboCity")))
            patient_province_field.click()
            search_province = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
            search_province.click()
            search_province.send_keys(patient_province)
            chooser_search_province = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
            chooser_search_province.click()
            #Nhập thành phố
            patient_ward_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboDistrict")))
            patient_ward_field.click()
            search_ward = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
            search_ward.click()
            search_ward.send_keys(patient_ward)
            chooser_search_ward = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
            chooser_search_ward.click()
            #Nhập phường
            patient_town_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboAuWard")))
            patient_town_field.click()
            search_town = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
            search_town.click()
            search_town.send_keys(patient_town)
            chooser_search_town = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
            chooser_search_town.click()
            #Nhập công ty
            patient_company_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "slueDvct")))
            patient_company_field.click()
            search_company = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
            search_company.click()
            search_company.send_keys(patient_company)
            chooser_search_company = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên đơn vị công tác row0")))
            chooser_search_company.click()
            #Nhập quyền lợi của BN
            patient_benefit_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboInsBenefitType")))
            patient_benefit_field.click()
            patient_benefit_field.clear()
            patient_benefit_field.send_keys(patient_benefit)
            chooser_search_benefit = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 0")))
            chooser_search_benefit.click()
            #Nếu quyền lợi là BHYT thì sẽ nhập thẻ BHYT, giá trị từ ngày đến ngày và nơi đăng ký
            if patient_benefit == "BHYT":
                #Nhập thẻ BHYT
                patient_bhyt_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtInsCardNo")))
                patient_bhyt_field.click()
                patient_bhyt_field.clear()
                patient_bhyt_field.send_keys(patient_bhyt)
                #Nhập nơi ĐK
                patient_place_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboMedProvider")))
                patient_place_field.click()
                search_place = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
                search_place.click()
                search_place.send_keys(patient_place)
                chooser_search_place = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
                chooser_search_place.click()
                #Giá trị từ ngày
                patient_from_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "dpInsValidFrom")))
                patient_from_field.click()
                formatted_from = pd.to_datetime(patient_from).strftime('%d%m%Y')
                patient_from_field.send_keys(formatted_from)
                #Giá trị đến ngày
                patient_to_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "dpInsValidTo")))
                patient_to_field.click()
                formatted_to = pd.to_datetime(patient_to).strftime('%d%m%Y')
                patient_to_field.send_keys(formatted_to)
            #Nhập loại khám
            if patient_benefit == "BHYT":
                patient_type_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboVisitReceiveType")))
                patient_type_field.click()
            else:
                patient_type_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboVisitReceiveType")))
                patient_type_field.click()
                patient_type_field.click()
            if patient_type == "Khám thường":
                chooser_patient_type = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 0")))
                chooser_patient_type.click()
            elif patient_type == "Khám cấp cứu":
                chooser_patient_type = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 1")))
                chooser_patient_type.click()
            elif patient_type == "Khám sức khỏe":
                chooser_patient_type = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 2")))
                chooser_patient_type.click()
            else:
                chooser_patient_type = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 3")))
                chooser_patient_type.click()
            #Nhập phòng khám
            patient_clinic_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboWardUnit")))
            patient_clinic_field.click()
            search_clinic = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
            search_clinic.click()
            if patient_clinic is None or isinstance(patient_clinic, float) and math.isnan(
                    patient_clinic):
                chooser_delete = WebDriverWait(self.appium_setup, 10).until(
                    EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btClear")))
                chooser_delete.click()
            else:
                search_clinic.send_keys(patient_clinic)
                chooser_search_clinic = WebDriverWait(self.appium_setup, 10).until(
                    EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
                chooser_search_clinic.click()
                if age < 6:
                    dialog = WebDriverWait(self.appium_setup, 10).until(
                        EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
                    dialog.click()
                    button_yes = WebDriverWait(self.appium_setup, 10).until(
                        EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "6")))
                    button_yes.click()
                # Nhập dịch vụ khám
                patient_service_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboMedService")))
                patient_service_field.click()
                search_service = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "teFind")))
                search_service.click()
                if patient_service is None or isinstance(patient_service, float) and math.isnan(
                        patient_service):
                    chooser_delete = WebDriverWait(self.appium_setup, 10).until(
                        EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btClear")))
                    chooser_delete.click()
                else:
                    search_service.send_keys(patient_service)
                    chooser_search_service = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Tên row0")))
                    chooser_search_service.click()
            #Click btn "Lưu" BN
            confirm_button = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnSave")))
            confirm_button.click()
            if age < 6 and patient_clinic is not None and patient_service is not None:
                dialog = WebDriverWait(self.appium_setup, 10).until(
                    EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
                dialog.click()
                button_yes = WebDriverWait(self.appium_setup, 10).until(
                    EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "6")))
                button_yes.click()
                time.sleep(5)
            # Kiểm tra xem cảnh báo đã xuất hiện hay không
            try:
                alert_element = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
                alert_element.click()
                print("Warning alert displayed.")
                if patient_benefit == "BHYT":
                    # Nhấn phím "Esc" trên bàn phím
                    pyautogui.press('esc')
                # Lưu mã bệnh nhân vào file datatest
                patient_code_field = WebDriverWait(self.appium_setup, 10).until(
                    EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtPatientCode")))
                patient_code_field.click()
                displayed_text = patient_code_field.text
                print(displayed_text)
                df.loc[df['TC_ID'] == TC_ID, 'patient_code'] = float(displayed_text)
                # Đọc dữ liệu hiện có từ tệp Excel
                existing_df = pd.read_excel(file_path)
                # Xóa các dòng có giá trị của cột 'TC_ID' là 'TC_3'
                existing_df = existing_df[existing_df['TC_ID'] != TC_ID]
                # Nối DataFrame hiện có và DataFrame mới (chỉ chứa dữ liệu của TC_ID là TC_3)
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                # Ghi toàn bộ DataFrame mới vào tệp Excel
                with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
                    combined_df.to_excel(writer, index=False, sheet_name='Sheet1')
                time.sleep(2)
            except:
                print("Warning alert not displayed.")
                if patient_benefit == "BHYT":
                    # Nhấn phím "Esc" trên bàn phím
                    pyautogui.press('esc')
                # Lưu mã bệnh nhân vào file datatest
                patient_code_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtPatientCode")))
                patient_code_field.click()
                displayed_text = patient_code_field.text
                print(displayed_text)
                df.loc[df['TC_ID'] == TC_ID, 'patient_code'] = float(displayed_text)
                # df.to_excel(file_path, index=False)
                # Đọc dữ liệu hiện có từ tệp Excel
                existing_df = pd.read_excel(file_path)
                # Xóa các dòng có giá trị của cột 'TC_ID' là 'TC_3'
                existing_df = existing_df[existing_df['TC_ID'] != TC_ID]
                # Nối DataFrame hiện có và DataFrame mới (chỉ chứa dữ liệu của TC_ID là TC_3)
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                # Ghi toàn bộ DataFrame mới vào tệp Excel
                with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
                    combined_df.to_excel(writer, index=False, sheet_name='Sheet1')
                time.sleep(2)
    def click_perform_preclinical(self):
        button_perform_preclinical = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnAddMedService")))
        button_perform_preclinical.click()
        try:
            perform_preclinical_screen = WebDriverWait(self.appium_setup, 10).until(
                EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "FrmDialog")))
            perform_preclinical_screen.click()
        except TimeoutException:
            inform_screen = WebDriverWait(self.appium_setup, 10).until(
                EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
            inform_screen.click()
            button_ok = WebDriverWait(self.appium_setup, 10).until(
                EC.presence_of_element_located((MobileBy.NAME, "&OK")))
            button_ok.click()
    def close_perform_preclinical(self):
        button_close = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnClose")))
        button_close.click()
    def printer_STT(self):
        button_printer = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnInSTTKham")))
        button_printer.click()
        printer_screen = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "FrmReportViewer")))
        printer_screen.click()
        pyautogui.press('esc')
    def printer_tem(self):
        button_printer_tem = WebDriverWait(self.appium_setup, 10).until(
            EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnInTem")))
        button_printer_tem.click()
        printer_tem_screen = WebDriverWait(self.appium_setup, 10).until(
            EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "FrmReportViewer")))
        printer_tem_screen.click()
        pyautogui.press('esc')
    def turn_off_inform(self):
        alert_element = WebDriverWait(self.appium_setup, 10).until(
            EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
        alert_element.click()
        button_ok = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "&OK")))
        button_ok.click()
        time.sleep(2)
    def close_application(self):
        button_close_app = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "Close")))
        button_close_app.click()
        inform = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
        inform.click()
        button_yes = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "&Yes")))
        button_yes.click()
    def update_information(self, TC_ID):
        file_path = 'D:/HIS automation/Datatest/datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
        # Lọc dữ liệu cho TC_ID được truyền vào
        df = df[df['TC_ID'] == TC_ID]
        # Lặp qua từng hàng trong DataFrame và nhập dữ liệu vào trường văn bản trên giao diện người dùng
        for index, row in df.iterrows():
            patient_code = row['patient_code']
            patient_bod = row['patient_bod']
            patient_email = row['patient_email']
            patient_phone = row['patient_phone']
            patient_blood = row['patient_blood']
            patient_blood_type = row['patient_blood_type']
            patient_cccd = row['patient_cccd']
            patient_cccd_date = row['patient_cccd_date']
            patient_cccd_place = row['patient_cccd_place']
            # Nhập mã bệnh nhân cũ
            patient_name_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtFullName")))
            patient_name_field.click()
            patient_name_field.send_keys(str(int(patient_code)))
            pyautogui.press('enter')
            dialog = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
            dialog.click()
            time.sleep(3)
            button_yes = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "6")))
            button_yes.click()
            age = function_visit.calculate_age(self, patient_bod)
            if age < 6:
                dialog = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Thông báo")))
                dialog.click()
                button_yes = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "6")))
                button_yes.click()
            # Click btn "Mã bệnh nhân"
            button_patient_code = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnPatientCode")))
            button_patient_code.click()
            time.sleep(2)
            # Nhập email
            patient_email_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtEmail")))
            patient_email_field.click()
            patient_email_field.send_keys(patient_email)
            # Nhập số điện thoại
            patient_phone_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtMobileNo")))
            patient_phone_field.click()
            patient_phone_field.send_keys("0" + str(patient_phone))
            # Nhập nhóm máu
            patient_blood_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboBloodTypeABO")))
            patient_blood_field.click()
            if patient_blood == "A":
                chooser_blood = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 0")))
                chooser_blood.click()
            elif patient_blood == "B":
                chooser_blood = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 1")))
                chooser_blood.click()
            elif patient_blood == "AB":
                chooser_blood = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 2")))
                chooser_blood.click()
            else:
                chooser_blood = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 3")))
                chooser_blood.click()
            patient_blood_type_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "cboBloodTypeRhd")))
            patient_blood_type_field.click()
            if patient_blood_type == "Dương":
                chooser_blood_type = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 0")))
                chooser_blood_type.click()
            else:
                chooser_blood_type = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.NAME, "Row 1")))
                chooser_blood_type.click()
            # Nhập căn cước công dân
            patient_cccd_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtIdCardNo")))
            patient_cccd_field.click()
            patient_cccd_field.send_keys("0" + str(patient_cccd))
            # Nhập ngày cấp cccd
            patient_cccd_date_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "deIdCardIssueOn")))
            patient_cccd_date_field.click()
            formatted_date = pd.to_datetime(patient_cccd_date).strftime('%d%m%Y')
            patient_cccd_date_field.send_keys(formatted_date)
            # Nhập nơi cấp cccd
            patient_cccd_place_field = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "txtIdCardIssueAt")))
            patient_cccd_place_field.click()
            patient_cccd_place_field.send_keys(patient_cccd_place)
            # Click btn Lưu
            button_save = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnSave")))
            button_save.click()
            # Click btn Đóng
            button_close = WebDriverWait(self.appium_setup, 10).until(EC.presence_of_element_located((MobileBy.ACCESSIBILITY_ID, "btnClose")))
            button_close.click()