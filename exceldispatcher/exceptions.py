
class ExcelDispatcherException(Exception):

    def __init__(self):
        self.message = "Excel Dispatcher Exception"


    def __str__(self) -> str:
        return self.message



class EngineFailedToStartException(ExcelDispatcherException):

    def __init__(self):
        self.message = "Excel engine failed to start"



class WorkbookException(ExcelDispatcherException):

    def __init__(self, workbook: str = ""):
        self.message = "Can't open workbook: " + workbook



class WorksheetException(ExcelDispatcherException):

    def __init__(self, worksheet: str = ""):
        self.message = "Can't open worksheet: " + worksheet



class FailedOnWriteException(ExcelDispatcherException):

    def __init__(self, row: int, column: int, value: str):
        self.message = "Failed to write value: " + value + " to row: " + str(row) + " column: " + str(column)



class NoWorkspaceException(ExcelDispatcherException):

    def __init__(self):
        self.message = "You must open a workbook and a worksheet before you can use it"


class FailedOnReadException(ExcelDispatcherException):

    def __init__(self, row: int, column: int):
        self.message = "Failed to read value from row: " + str(row) + " column: " + str(column)



class FailedOnSaveException(ExcelDispatcherException):

    def __init__(self):
        self.message = "Something goes wrong with the save operation"