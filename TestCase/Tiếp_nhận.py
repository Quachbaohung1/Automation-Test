import unittest
import sys
sys.path.append('D:/HIS automation')
from Driver.init_driver import AppiumSetup
from Phân_hệ.Tiếp_nhận_Test import function_visit

class CustomTestResult(unittest.TestResult):
    def __init__(self, stream=None, descriptions=None, verbosity=None):
        super().__init__(stream, descriptions, verbosity)
        self.results = []

    def addSuccess(self, test):
        super().addSuccess(test)
        self.results.append((test, "Passed"))

    def addFailure(self, test, err):
        super().addFailure(test, err)
        self.results.append((test, "Failed"))

    def addError(self, test, err):
        super().addError(test, err)
        self.results.append((test, "Error"))

    def get_results(self):
        return self.results

    def get_passed_results(self):
        return [result for result in self.results if result[1] == "Passed"]

class VisitTestSetup(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.appium_setup = AppiumSetup()
        cls.appium_setup.start_appium_session()

    @classmethod
    def tearDownClass(cls):
        cls.appium_setup.quit_appium_session()

    # Tiếp nhận thiếu field "Họ tên"
    def test_case_1(self):
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_1")
        function_visit.turn_off_inform(self)
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Tiếp nhận thiếu field "Ngày sinh"
    def test_case_2(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_2")
        function_visit.turn_off_inform(self)
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Tiếp nhận thiếu field "Quốc tịch"
    def test_case_3(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_3")
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Tiếp nhận thiếu field "Số nhà/ đường"
    def test_case_4(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_4")
        function_visit.turn_off_inform(self)
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Tiếp nhận thiếu field "Phòng khám"
    def test_case_5(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_5")
        function_visit.turn_off_inform(self)
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Tiếp nhận thiếu field "Dịch vụ khám"
    def test_case_6(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_6")
        function_visit.turn_off_inform(self)
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Tiếp nhận thành công
    def test_case_7(self):
        # self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_7")
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Bổ sung thông tin thành công
    def test_case_8(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.update_information(self, TC_ID = "TC_7")
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Click btn " Chỉ định dịch vụ" -> Nếu là Thu phí thì có hiển thị, BHYT thì không hiển thị
    def test_case_9(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_9")
        function_visit.click_perform_preclinical(self)
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Click btn "In STT khám"
    def test_case_10(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_10")
        function_visit.printer_STT(self)
        function_visit.close_application(self)
        self.appium_setup.quit_appium_session()
    # Click btn "In tem"
    def test_case_11(self):
        self.appium_setup.start_appium_session()
        function_visit.login_visit(self)
        function_visit.visit_function(self, TC_ID = "TC_11")
        function_visit.printer_tem(self)
        function_visit.close_application(self)
        # self.appium_setup.quit_appium_session()
    def generate_test_report(self):
        print("Generating test report...")
        # Tạo một TestSuite từ lớp VisitTestSetup
        suite = unittest.TestLoader().loadTestsFromTestCase(self.__class__)
        # Chạy TestSuite và lấy kết quả
        custom_result = CustomTestResult()
        suite.run(custom_result)

        successes = custom_result.testsRun - len(custom_result.failures) - len(custom_result.errors)

        # Ghi kết quả vào file
        with open("D:/HIS automation/Result/TestResult.txt", "w", encoding="utf-8") as file:
            file.write("=== Test Report ===\n")
            file.write(f"Total tests run: {custom_result.testsRun}\n")
            file.write(f"Successes: {successes}\n")
            file.write(f"Failures: {len(custom_result.failures)}\n")
            file.write(f"Errors: {len(custom_result.errors)}\n")
            file.write(f"Skipped: {len(custom_result.skipped)}\n")

            if successes > 0:
                file.write("\n=== Successes ===\n")
                for test, _ in custom_result.get_passed_results():
                    file.write(f"{test}: Passed\n")

            if custom_result.failures:
                file.write("\n=== Failures ===\n")
                for test, fail_msg in custom_result.failures:
                    file.write(f"{test}: {fail_msg}\n")

            if custom_result.errors:
                file.write("\n=== Errors ===\n")
                for test, err_msg in custom_result.errors:
                    file.write(f"{test}: {err_msg}\n")
        print("Test report generated.")
        return custom_result
def main():
    test_setup = VisitTestSetup()
    test_setup.generate_test_report()

if __name__ == "__main__":
    main()