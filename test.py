from exceldispatcher.exceldispatcher import ExcelDispatcher
from fake_const import *

REPORT_TEST_FILE = "D:\\ocsubsystem\\Report_Numerico_INBOUND_BASE.xlsx"
REPORT_TEST_WORKSHEET = "Nome_attività"


# excel dispatcher test
try:
    excel = ExcelDispatcher(workbook=REPORT_TEST_FILE)
except Exception as e:
    print(e)


# worksheet opened except: ExcelDispatcher Exception

try:
    excel.open_worksheet(REPORT_TEST_WORKSHEET)
except Exception as e:
    print(e)


# try to get the first report date value
print(excel.read(12, 4))





excel.save_and_quit()