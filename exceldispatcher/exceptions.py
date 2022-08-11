class ExcelDispatcherException(Exception):
    def __init__(self):
        self.message = "Excel Dispatcher Exception"

    def __str__(self) -> str:
        return self.message


class EngineFailedToStartException(ExcelDispatcherException):
    def __init__(self):
        self.message = "Excel engine failed to start"



class WorkbookAndWorksheetException(ExcelDispatcherException):
    def __init(self):
        self.message = "Something goes wrong with the workbook or with the worksheet"



class WorkbookException(ExcelDispatcherException):
    def __init__(self):
        self.message = "Something goes wrong with the workbook"


class worksheetException(ExcelDispatcherException):
    def __init__(self):
        self.message = "Something goes wrong with the worksheet"


class FailedOnWriteException(ExcelDispatcherException):
    def __init__(self):
        self.message = "Failed to write on file."

class FailedOnReadException(ExcelDispatcherException):
    def __init__(self):
        self.message = "Failed to read cell from file."



class FileException(ExcelDispatcherException):
    def __init__(self):
        self.message = "Something goes wrong with the file infos."