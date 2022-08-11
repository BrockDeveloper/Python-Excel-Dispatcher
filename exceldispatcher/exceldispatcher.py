from typing import Any
import win32com.client as win32
from exceldispatcher.exceptions import *


CALLABLE = "Excel.Application"


class ExcelDispatcher:



    def __init__(self, workbook: str = None, worksheet: str = None) -> None:

        '''
        Create a new excel engine, optionally open a workbook and a worksheet

        Parameters:
        - workbook (str): Path to the workbook to open
        - worksheet (str): Name of the worksheet to open

        Returns:
            None
        '''

        self.workbook = None
        self.worksheet = None

        # Create a new instance of Excel [engine]
        try:
            self.engine = win32.gencache.EnsureDispatch(CALLABLE)
        except Exception:
            raise EngineFailedToStartException()

        # headless excel session
        self.engine.DisplayAlerts = False                                   
        self.engine.Visible = False                                        
        self.engine.ScreenUpdating = False

        # Open the workbook, if specified
        if workbook:
            try:
                self.workbook = self.engine.Workbooks.Open(workbook)
            except Exception:
                raise WorkbookException()

            # Open the worksheet, if specified
            if worksheet:
                try:
                    self.open_worksheet(worksheet)
                except Exception:
                    raise worksheetException()



    def open_workbook(self, file_path: str) -> None:
        try:
            self.workbook = self.engine.Workbooks.Open(file_path)
        except Exception:
            raise WorkbookException()



    def open_worksheet(self, worksheet: str) -> None:
        try:
            self.worksheet = self.workbook.Worksheets(worksheet)
        except Exception:
            raise worksheetException()





    # get the current workbook name
    def get_workbook_name(self) -> str:
        if self.workbook:
            return self.workbook.Name
        else:
            return ""

    # get the current workbook path
    def get_workbook_dir(self) -> str:
        if self.workbook:
            return self.workbook.Path
        else:
            return ""

    # get the current worksheet complete path with name
    def get_workbook_path(self) -> str:
        if self.workbook and self.worksheet:
            return self.workbook.Path + "\\" + self.workbook.Name
        else:
            return ""


    # get the current worksheet name
    def get_worksheet_name(self) -> str:
        if self.worksheet:
            return self.worksheet.Name
        else:
            return ""


    def list_worksheets(self) -> list:
        if self.workbook:
            return [sheet.Name for sheet in self.workbook.Worksheets]
        else:
            return []


    def close_and_quit(self):
        self.engine.ScreenUpdating = True
        self.engine.Application.Quit()


    def write(self, value: Any, row: int, col: int) -> None:
        try:
            self.worksheet.Cells(row, col).Value = value
        except Exception as e:
            raise FailedOnWriteException()


    def read(self, row: int, col: int) -> Any:
        try:
            return self.worksheet.Cells(row, col).Value
        except Exception as e:
            raise FailedOnReadException()


    def save(self) -> None:
        try:
            self.workbook.Save()
        except Exception as e:
            raise ExcelDispatcherException()

    
    def save_and_quit(self) -> None:
        try:
            self.save()
            self.close_and_quit()
        except Exception as e:
            raise ExcelDispatcherException()
           

    # call if except occurred
    def __close_and_quit_on_error(self, error: Exception) -> None:
        if "engine" in locals():
            self.engine.ScreenUpdating = True
            self.engine.Application.Quit()
        else:
            raise ExcelDispatcherException()
    



    # get the number of used rows
    def get_used_rows(self) -> int:
        
        # check if workbook and worksheet are definied
        if self.workbook and self.worksheet:
            return self.worksheet.UsedRange.Rows.Count
        else:
            raise WorkbookAndWorksheetException()

        

    def get_used_cols(self) -> int:        

        '''
        Return the number of the used columns

        PARAM:
            None

        RETURN:
            int, the number of the cols
        '''

        # check if workbook and worksheet are definied
        if self.workbook and self.worksheet:
            return self.worksheet.UsedRange.Columns.Count
        else:
            raise WorkbookAndWorksheetException()





    
    # fit column width to fit the content
    def fit_column_width(self, col: int) -> None:
        self.worksheet.Columns.AutoFit()