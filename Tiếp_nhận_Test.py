import unittest
import datetime
import pyautogui
from init_driver import AppiumSetup
import time
import pandas as pd


class function():
    def __init__(self):
        self.appium_setup = AppiumSetup()

    def calculate_age(self, patient_bod):
        # Chuyển đổi ngày tháng năm sinh từ định dạng chuỗi sang đối tượng datetime
        patient_bod_datetime = pd.to_datetime(patient_bod, format='%Y-%m-%d')
        today = datetime.datetime.today()
        age = today.year - patient_bod_datetime.year - (
                (today.month, today.day) < (patient_bod_datetime.month, patient_bod_datetime.day))
        return age

    def login_with_parameters(self):
        # Tiếp tục làm việc với ứng dụng của bạn bằng cách sử dụng các lớp cửa sổ/phương thức tương ứng
        time.sleep(3)
        element = self.appium_setup.find_element_by_accessibility_id("FrmMain")
        element.click()

        # Nhập Tài khoản
        account_field = self.appium_setup.find_element_by_accessibility_id("txtAccName")
        account_field.click()  # Click vào trường "Tài khoản" để đảm bảo trường đóng được chọn
        # Xóa nội dung trường "Tài khoản" trước khi nhập liệu mới
        account_field.clear()
        # Nhập dữ liệu mới vào trường "Tài khoản"
        account_field.send_keys("hungqb")

        # Nhập Mật khẩu
        password_field = self.appium_setup.find_element_by_accessibility_id("txtPassword")
        password_field.click()  # Click vào trường "Tài khoản" để đảm bảo trường đóng được chọn
        # Xóa nội dung trường "Mật khẩu" trước khi nhập liệu mới
        password_field.clear()
        # Nhập dữ liệu mới vào trường "Tài khoản"
        password_field.send_keys("1")

        # Click btn Đăng nhập
        login = self.appium_setup.find_element_by_accessibility_id("btnLogin")
        login.click()

class VisitTestSetup(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.appium_setup = AppiumSetup()
        cls.appium_setup.start_appium_session()

    @classmethod
    def tearDownClass(cls):
        cls.appium_setup.quit_appium_session()
    def test_login(self):
        function.login_with_parameters(self)
        # Click chọn Khoa
        department = self.appium_setup.find_element_by_accessibility_id("cboWards")
        self.assertIsNotNone(department, "Không tìm thấy phần tử bằng accessibility_id")
        department.click()
        search_department = self.appium_setup.find_element_by_accessibility_id("teFind")
        self.assertIsNotNone(search_department, "Không tìm thấy phần tử bằng accessibility_id")
        search_department.click()
        search_department.send_keys("Phòng điều dưỡng")
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
        search_room.send_keys("Tiếp nhận khu công")
        chooser_search_room = self.appium_setup.find_element_by_name("gridColumn2 row0")
        self.assertIsNotNone(chooser_search_room, "Không tìm thấy phần tử bằng name")
        chooser_search_room.click()

        # Click btn Xác nhận
        confirmed = self.appium_setup.find_element_by_accessibility_id("btnLogin")
        self.assertIsNotNone(confirmed, "Không tìm thấy phần tử bằng accessibility_id")
        confirmed.click()

        #Click chức năng Tiếp nhận
        visit_click = self.appium_setup.find_element_by_accessibility_id("btnVisit")
        self.assertIsNotNone(visit_click, "Không tìm thấy phần tử bằng accessibility_id")
        visit_click.click()
        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(5)

        #Mở màn hình tiếp nhận
        visit_menu = self.appium_setup.find_element_by_name("Tiếp nhận")
        self.assertIsNotNone(visit_menu, "Không tìm thấy phần tử bằng name")
        visit_menu.click()
        visit_item = self.appium_setup.find_element_by_xpath("//Button[@Name='Tiếp nhận']")
        self.assertIsNotNone(visit_item, "Không tìm thấy phần tử bằng xpath")
        visit_item.click()
        # Đợi một khoảng thời gian ngắn để ứng dụng mở.
        time.sleep(5)

    def test_reception_function(self):
        #Đọc file data test
        file_path = 'datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
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
            element_visit = self.appium_setup.find_element_by_accessibility_id("pnMainPage")
            element_visit.click()
            self.assertIsNotNone(element_visit, "Không tìm thấy phần tử bằng accessibility_id")

            #Thực hiện các bước để tiếp nhận bệnh nhân với tên và mã bệnh nhân từ file XLSX
            #Nhập họ và tên
            patient_name_field = self.appium_setup.find_element_by_accessibility_id("txtFullName")
            patient_name_field.click()
            patient_name_field.send_keys(patient_name)

            #Nhập giới tính
            patient_gender_field = self.appium_setup.find_element_by_accessibility_id("cboGender")
            patient_gender_field.click()
            patient_gender_field.send_keys(patient_gender)

            #Nhập ngày sinh
            patient_bod_field = self.appium_setup.find_element_by_accessibility_id("dpDOB")
            patient_bod_field.click()
            # Chuyển đổi ngày sinh sang định dạng mong muốn
            formatted_bod = pd.to_datetime(patient_bod).strftime('%d%m%Y')
            patient_bod_field.send_keys(formatted_bod)

            # Tính toán tuổi của bệnh nhân
            age = VisitTestSetup.calculate_age(self, patient_bod)
            # Kiểm tra xem tuổi của bệnh nhân có dưới 6 không
            if age < 6:
                print("Tuổi của bệnh nhân bé hơn 6 tuổi. Cần có người thân.")
                #Nhập mối quan hệ với bệnh nhân
                patient_relative_field = self.appium_setup.find_element_by_accessibility_id("cboRelativeType")
                patient_relative_field.click()
                if patient_relative == "Cha":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 1")
                    chooser_patient_relative.click()
                elif patient_relative == "Mẹ":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 2")
                    chooser_patient_relative.click()
                elif patient_relative == "Anh em ruột":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 3")
                    chooser_patient_relative.click()
                elif patient_relative == "Con":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 4")
                    chooser_patient_relative.click()
                elif patient_relative == "Người bảo hộ":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 5")
                    chooser_patient_relative.click()
                elif patient_relative == "Vợ-Chồng":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 6")
                    chooser_patient_relative.click()
                elif patient_relative == "Người đưa đến":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 7")
                    chooser_patient_relative.click()
                elif patient_relative == "Ông-Bà":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 8")
                    chooser_patient_relative.click()
                elif patient_relative == "Cháu":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 9")
                    chooser_patient_relative.click()
                elif patient_relative == "Cô":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 10")
                    chooser_patient_relative.click()
                elif patient_relative == "Anh":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 11")
                    chooser_patient_relative.click()
                elif patient_relative == "Em":
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 12")
                    chooser_patient_relative.click()
                else:
                    chooser_patient_relative = self.appium_setup.find_element_by_name("Row 13")
                    chooser_patient_relative.click()

                #Nhập tên người thân
                patient_relativeName_field = self.appium_setup.find_element_by_accessibility_id("txtRelativeName")
                patient_relativeName_field.click()
                patient_relativeName_field.send_keys(patient_relativeName)

                #Nhập số điện thoại người thân
                patient_relativePhone_field = self.appium_setup.find_element_by_accessibility_id("txtRelativePhone")
                patient_relativePhone_field.click()
                patient_relativePhone_field.send_keys("0" + str(patient_relativePhone))

                #Nhập địa chỉ của người thân
                patient_relativeAddress_field = self.appium_setup.find_element_by_accessibility_id("txtRelativeAddress")
                patient_relativeAddress_field.click()
                patient_relativeAddress_field.send_keys(patient_relativeAddress)
            else:
                print("Tuổi của bệnh nhân lớn hơn hoặc bằng 6 tuổi. Không cần có người thân.")

            #Nhập quốc tịch
            patient_nationality_field = self.appium_setup.find_element_by_accessibility_id("cboNationary")
            patient_nationality_field.click()
            search_nationality = self.appium_setup.find_element_by_accessibility_id("teFind")
            self.assertIsNotNone(search_nationality, "Không tìm thấy phần tử bằng accessibility_id")
            search_nationality.click()
            search_nationality.send_keys(patient_nationality)
            chooser_search_nationality = self.appium_setup.find_element_by_name("Tên row0")
            self.assertIsNotNone(chooser_search_nationality, "Không tìm thấy phần tử bằng name")
            chooser_search_nationality.click()

            #Nếu quốc tịch là VN thì sẽ nhập dân tộc, nếu khác VN thì không cần nhập ở field này
            if patient_nationality == "Việt Nam":
                patient_nation_field = self.appium_setup.find_element_by_accessibility_id("cboEthnic")
                patient_nation_field.click()
                patient_nation_field.clear()
                patient_nation_field.send_keys(patient_nation)

            #Nhập nghề nghiệp
            patient_job_field = self.appium_setup.find_element_by_accessibility_id("cboOccupation")
            patient_job_field.click()
            patient_job_field.clear()
            patient_job_field.send_keys(patient_job)

            #Nhập địa chỉ
            patient_address_field = self.appium_setup.find_element_by_accessibility_id("txtAddress")
            patient_address_field.click()
            patient_address_field.send_keys(patient_address)

            #Nhập tỉnh
            patient_province_field = self.appium_setup.find_element_by_accessibility_id("cboCity")
            patient_province_field.click()
            search_province = self.appium_setup.find_element_by_accessibility_id("teFind")
            self.assertIsNotNone(search_province, "Không tìm thấy phần tử bằng accessibility_id")
            search_province.click()
            search_province.send_keys(patient_province)
            chooser_search_province = self.appium_setup.find_element_by_name("Tên row0")
            self.assertIsNotNone(chooser_search_province, "Không tìm thấy phần tử bằng name")
            chooser_search_province.click()

            #Nhập thành phố
            patient_ward_field = self.appium_setup.find_element_by_accessibility_id("cboDistrict")
            patient_ward_field.click()
            search_ward = self.appium_setup.find_element_by_accessibility_id("teFind")
            self.assertIsNotNone(search_ward, "Không tìm thấy phần tử bằng accessibility_id")
            search_ward.click()
            search_ward.send_keys(patient_ward)
            chooser_search_ward = self.appium_setup.find_element_by_name("Tên row0")
            self.assertIsNotNone(chooser_search_ward, "Không tìm thấy phần tử bằng name")
            chooser_search_ward.click()

            #Nhập phường
            patient_town_field = self.appium_setup.find_element_by_accessibility_id("cboAuWard")
            patient_town_field.click()
            search_town = self.appium_setup.find_element_by_accessibility_id("teFind")
            self.assertIsNotNone(search_town, "Không tìm thấy phần tử bằng accessibility_id")
            search_town.click()
            search_town.send_keys(patient_town)
            chooser_search_town = self.appium_setup.find_element_by_name("Tên row0")
            self.assertIsNotNone(chooser_search_town, "Không tìm thấy phần tử bằng name")
            chooser_search_town.click()

            #Nhập công ty
            patient_company_field = self.appium_setup.find_element_by_accessibility_id("slueDvct")
            patient_company_field.click()
            search_company = self.appium_setup.find_element_by_accessibility_id("teFind")
            self.assertIsNotNone(search_company, "Không tìm thấy phần tử bằng accessibility_id")
            search_company.click()
            search_company.send_keys(patient_company)
            chooser_search_company = self.appium_setup.find_element_by_name("Tên đơn vị công tác row0")
            self.assertIsNotNone(chooser_search_company, "Không tìm thấy phần tử bằng name")
            chooser_search_company.click()

            #Nhập quyền lợi của BN
            patient_benefit_field = self.appium_setup.find_element_by_accessibility_id("cboInsBenefitType")
            patient_benefit_field.click()
            patient_benefit_field.clear()
            patient_benefit_field.send_keys(patient_benefit)
            chooser_search_benefit = self.appium_setup.find_element_by_name("Row 0")
            self.assertIsNotNone(chooser_search_benefit, "Không tìm thấy phần tử bằng name")
            chooser_search_benefit.click()

            #Nếu quyền lợi là BHYT thì sẽ nhập thẻ BHYT, giá trị từ ngày đến ngày và nơi đăng ký
            if patient_benefit == "BHYT":
                #Nhập thẻ BHYT
                patient_bhyt_field = self.appium_setup.find_element_by_accessibility_id("txtInsCardNo")
                patient_bhyt_field.click()
                patient_bhyt_field.clear()
                patient_bhyt_field.send_keys(patient_bhyt)

                #Nhập nơi ĐK
                patient_place_field = self.appium_setup.find_element_by_accessibility_id("cboMedProvider")
                patient_place_field.click()
                search_place = self.appium_setup.find_element_by_accessibility_id("teFind")
                self.assertIsNotNone(search_place, "Không tìm thấy phần tử bằng accessibility_id")
                search_place.click()
                search_place.send_keys(patient_place)
                chooser_search_place = self.appium_setup.find_element_by_name("Tên row0")
                self.assertIsNotNone(chooser_search_place, "Không tìm thấy phần tử bằng name")
                chooser_search_place.click()

                #Giá trị từ ngày
                patient_from_field = self.appium_setup.find_element_by_accessibility_id("dpInsValidFrom")
                patient_from_field.click()
                formatted_from = pd.to_datetime(patient_from).strftime('%d%m%Y')
                patient_from_field.send_keys(formatted_from)

                #Giá trị đến ngày
                patient_to_field = self.appium_setup.find_element_by_accessibility_id("dpInsValidTo")
                patient_to_field.click()
                formatted_to = pd.to_datetime(patient_to).strftime('%d%m%Y')
                patient_to_field.send_keys(formatted_to)

            #Nhập loại khám
            if patient_benefit == "BHYT":
                patient_type_field = self.appium_setup.find_element_by_accessibility_id("cboVisitReceiveType")
                patient_type_field.click()
            else:
                patient_type_field = self.appium_setup.find_element_by_accessibility_id("cboVisitReceiveType")
                patient_type_field.click()
                patient_type_field.click()

            if patient_type == "Khám thường":
                chooser_patient_type = self.appium_setup.find_element_by_name("Row 0")
                chooser_patient_type.click()
            elif patient_type == "Khám cấp cứu":
                chooser_patient_type = self.appium_setup.find_element_by_name("Row 1")
                chooser_patient_type.click()
            elif patient_type == "Khám sức khỏe":
                chooser_patient_type = self.appium_setup.find_element_by_name("Row 2")
                chooser_patient_type.click()
            else:
                chooser_patient_type = self.appium_setup.find_element_by_name("Row 3")
                chooser_patient_type.click()

            #Nhập phòng khám
            patient_clinic_field = self.appium_setup.find_element_by_accessibility_id("cboWardUnit")
            patient_clinic_field.click()
            search_clinic = self.appium_setup.find_element_by_accessibility_id("teFind")
            self.assertIsNotNone(search_clinic, "Không tìm thấy phần tử bằng accessibility_id")
            search_clinic.click()
            search_clinic.send_keys(patient_clinic)
            chooser_search_clinic = self.appium_setup.find_element_by_name("Tên row0")
            self.assertIsNotNone(chooser_search_clinic, "Không tìm thấy phần tử bằng name")
            chooser_search_clinic.click()

            if age < 6:
                dialog = self.appium_setup.find_element_by_name("Thông báo")
                dialog.click()
                button_yes = self.appium_setup.find_element_by_accessibility_id("6")
                button_yes.click()
                time.sleep(10)

            #Nhập dịch vụ khám
            patient_service_field = self.appium_setup.find_element_by_accessibility_id("cboMedService")
            patient_service_field.click()
            search_service = self.appium_setup.find_element_by_accessibility_id("teFind")
            self.assertIsNotNone(search_service, "Không tìm thấy phần tử bằng accessibility_id")
            search_service.click()
            search_service.send_keys(patient_service)
            chooser_search_service = self.appium_setup.find_element_by_name("Tên row0")
            self.assertIsNotNone(chooser_search_service, "Không tìm thấy phần tử bằng name")
            chooser_search_service.click()

            #Click btn "Lưu" BN
            confirm_button = self.appium_setup.find_element_by_accessibility_id("btnSave")
            confirm_button.click()
            time.sleep(5)

            if age < 6:
                dialog = self.appium_setup.find_element_by_name("Thông báo")
                dialog.click()
                button_yes = self.appium_setup.find_element_by_accessibility_id("6")
                button_yes.click()
                time.sleep(10)

            if patient_benefit == "BHYT":
                # Nhấn phím "Esc" trên bàn phím
                pyautogui.press('esc')
                time.sleep(10)

            # Lưu mã bệnh nhân vào file datatest
            patient_code_field = self.appium_setup.find_element_by_accessibility_id("txtPatientCode")
            patient_code_field.click()
            displayed_text = patient_code_field.text
            print(displayed_text)
            df.at[index, 'patient_code'] = displayed_text
            df.to_excel(file_path, index=False)
            time.sleep(10)

            #Click btn "Làm mới" để tiếp nhận tiếp các bệnh nhân khác trong file datatest
            renew_button = self.appium_setup.find_element_by_accessibility_id("btnReNew")
            renew_button.click()
            time.sleep(5)

    def test_update_information(self):
        file_path = 'datatest.xlsx'
        # Đọc dữ liệu từ file Excel và lưu vào một DataFrame pandas
        df = pd.read_excel(file_path)
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

            # Lưu lại cửa sổ ban đầu
            original_window_handle = self.appium_setup.get_driver().current_window_handle

            #Nhập mã bệnh nhân cũ
            patient_name_field = self.appium_setup.find_element_by_accessibility_id("txtFullName")
            patient_name_field.click()
            patient_name_field.send_keys(patient_code)

            pyautogui.press('enter')
            time.sleep(10)

            dialog = self.appium_setup.find_element_by_name("Thông báo")
            dialog.click()
            button_yes = self.appium_setup.find_element_by_accessibility_id("6")
            button_yes.click()
            time.sleep(5)

            age = VisitTestSetup.calculate_age(self, patient_bod)
            if age < 6:
                dialog = self.appium_setup.find_element_by_name("Thông báo")
                dialog.click()
                button_yes = self.appium_setup.find_element_by_accessibility_id("6")
                button_yes.click()
                time.sleep(5)

            #Click btn "Mã bệnh nhân"
            button_patient_code = self.appium_setup.find_element_by_accessibility_id("btnPatientCode")
            button_patient_code.click()

            # Lấy tất cả các handles của cửa sổ
            all_windows = self.appium_setup.get_driver().window_handles
            # Chuyển đến cửa sổ mới dựa trên mã cửa sổ
            for window_handle in all_windows:
                # Kiểm tra xem mã cửa sổ có tồn tại trong danh sách hay không
                if window_handle != original_window_handle:
                    # Chuyển đến cửa sổ mới
                    self.appium_setup.get_driver().switch_to.window(window_handle)
                    time.sleep(3)
                    break
            # Thực hiện các thao tác trên cửa sổ mới
            #Nhập email
            patient_email_field = self.appium_setup.find_element_by_accessibility_id("txtEmail")
            patient_email_field.click()
            patient_email_field.send_keys(patient_email)

            #Nhập số điện thoại
            patient_phone_field = self.appium_setup.find_element_by_accessibility_id("txtMobileNo")
            patient_phone_field.click()
            patient_phone_field.send_keys("0" + str(patient_phone))

            #Nhập nhóm máu
            patient_blood_field = self.appium_setup.find_element_by_accessibility_id("cboBloodTypeABO")
            patient_blood_field.click()
            if patient_blood == "A":
                chooser_blood = self.appium_setup.find_element_by_name("Row 0")
                chooser_blood.click()
            elif patient_blood == "B":
                chooser_blood = self.appium_setup.find_element_by_name("Row 1")
                chooser_blood.click()
            elif patient_blood == "AB":
                chooser_blood = self.appium_setup.find_element_by_name("Row 2")
                chooser_blood.click()
            else:
                chooser_blood = self.appium_setup.find_element_by_name("Row 3")
                chooser_blood.click()

            patient_blood_type_field = self.appium_setup.find_element_by_accessibility_id("cboBloodTypeRhd")
            patient_blood_type_field.click()
            if patient_blood_type == "Dương":
                chooser_blood_type = self.appium_setup.find_element_by_name("Row 0")
                chooser_blood_type.click()
            else:
                chooser_blood_type = self.appium_setup.find_element_by_name("Row 1")
                chooser_blood_type.click()
            #Nhập căn cước công dân
            patient_cccd_field = self.appium_setup.find_element_by_accessibility_id("txtIdCardNo")
            patient_cccd_field.click()
            patient_cccd_field.send_keys("0" + str(patient_cccd))

            #Nhập ngày cấp cccd
            patient_cccd_date_field = self.appium_setup.find_element_by_accessibility_id("deIdCardIssueOn")
            patient_cccd_date_field.click()
            formatted_date = pd.to_datetime(patient_cccd_date).strftime('%d%m%Y')
            patient_cccd_date_field.send_keys(formatted_date)

            #Nhập nơi cấp cccd
            patient_cccd_place_field = self.appium_setup.find_element_by_accessibility_id("txtIdCardIssueAt")
            patient_cccd_place_field.click()
            patient_cccd_place_field.send_keys(patient_cccd_place)

            #Click btn Lưu
            button_save = self.appium_setup.find_element_by_accessibility_id("btnSave")
            button_save.click()

            #Click btn Đóng
            button_close = self.appium_setup.find_element_by_accessibility_id("btnClose")
            button_close.click()
            time.sleep(2)

            # Quay lại cửa sổ ban đầu
            self.appium_setup.get_driver().switch_to.window(original_window_handle)

if __name__ == "__main__":
    unittest.main()
