
from typing import Any
import win32com.client as win32
from exceldispatcher.exceptions import *


CALLABLE = "Excel.Application"


class ExcelDispatcher:



    def __init__(self, workbook: str = None, worksheet: str = None) -> None:

        '''
        Create a new excel engine, optionally open a workbook and a worksheet

        Parameters:
            workbook (str): Path to the workbook to open
            worksheet (str): Name of the worksheet to open

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
                raise WorkbookException(workbook)

            # Open the worksheet, if specified
            if worksheet:
                try:
                    self.open_worksheet(worksheet)
                except Exception:
                    raise WorksheetException(worksheet)



    def open_workbook(self, file_path: str) -> None:

        '''
        Open a workbook
        
        Parameters:
            file_path (str): Path to the workbook to open
        
        Returns:
            None
            
        Raises:
            WorkbookException
        '''

        try:
            self.workbook = self.engine.Workbooks.Open(file_path)
        except Exception:
            raise WorkbookException(file_path)



    def open_worksheet(self, worksheet: str) -> None:

        '''
        Open a worksheet
        
        Parameters:
            worksheet (str): Name of the worksheet to open
        
        Returns:
            None
            
        Raises:
            WorksheetException'''

        try:
            self.worksheet = self.workbook.Worksheets(worksheet)
        except Exception:
            raise WorksheetException(worksheet)



    def get_workbook_name(self) -> str:
        
        '''
        Return the name of the workbook
        
        Parameters:
            None
            
        Returns:
            the name of the workbook, if exists
            empty string, otherwise
        '''

        if self.workbook:
            return self.workbook.Name
        else:
            return ""



    def get_workbook_dir(self) -> str:

        '''
        Return the directory of the workbook
        
        Parameters:
            None
            
        Returns:
            the directory of the workbook, if exists
            empty string, otherwise
        '''

        if self.workbook:
            return self.workbook.Path
        else:
            return ""



    def get_workbook_path(self) -> str:

        '''
        Return the path of the workbook
        
        Parameters:
            None
            
        Returns:
            the path of the workbook, if exists
        '''

        if self.workbook and self.worksheet:
            return self.workbook.Path + "\\" + self.workbook.Name
        else:
            return ""



    def get_worksheet_name(self) -> str:

        '''
        Return the name of the worksheet
        
        Parameters:
            None
            
        Returns:
            the name of the worksheet, if exists
            empty string, otherwise
        '''

        if self.worksheet:
            return self.worksheet.Name
        else:
            return ""



    def list_worksheets(self) -> list:

        '''
        Return a list of the worksheets in the workbook
        
        Parameters:
            None
            
        Returns:
            a list of the worksheets in the workbook, if exists
            empty list, otherwise
        '''

        if self.workbook:
            return [sheet.Name for sheet in self.workbook.Worksheets]
        else:
            return []



    def write(self, value: Any, row: int, col: int) -> None:

        '''
        Write a value to a cell
        
        Parameters:
            value (Any): Value to write to the cell
            row (int): Row of the cell to write to
            col (int): Column of the cell to write to
        
        Returns:
            None
            
        Raises:
            FailedOnWriteException
        '''

        try:
            self.worksheet.Cells(row, col).Value = value
        except Exception as e:
            raise FailedOnWriteException(row, col, value)



    def read(self, row: int, col: int) -> Any:

        '''
        Read a value from a cell
        
        Parameters:
            row (int): Row of the cell to read from
            col (int): Column of the cell to read from
            
        Returns:
            the value of the cell, if exists
        
        Raises:
            FailedOnReadException
        '''

        try:
            return self.worksheet.Cells(row, col).Value
        except Exception as e:
            raise FailedOnReadException(row, col)



    def get_used_rows(self) -> int:

        '''
        Return the number of used rows in the worksheet

        Parameters:
            None

        Returns:
            the number of used rows in the worksheet, if exists

        Raises:
            FailedOnReadException
        '''
        
        if self.workbook and self.worksheet:
            return self.worksheet.UsedRange.Rows.Count
        else:
            raise NoWorkspaceException()

        

    def get_used_cols(self) -> int:        

        '''
        Return the number of used columns in the worksheet
        
        Parameters:
            None

        Returns:
            the number of used columns in the worksheet, if exists

        Raises:
            FailedOnReadException
        '''

        if self.workbook and self.worksheet:
            return self.worksheet.UsedRange.Columns.Count
        else:
            raise NoWorkspaceException()



    def fit_column_width(self, col: int) -> None:

        '''
        Fit the width of a column to its contents

        Parameters:
            col (int): Column to fit the width of

        Returns:
            None
        '''

        if self.workbook and self.worksheet:
            self.worksheet.Columns.AutoFit()
        else:
            raise NoWorkspaceException()



    def save(self) -> None:

        '''
        Save the workbook
        
        Parameters:
            None
            
        Returns:
            None
            
        Raises:
            FailedOnSaveException
        '''

        try:
            self.workbook.Save()
        except Exception as e:
            raise FailedOnSaveException()

    
    
    def close_and_quit(self) -> None:

        '''
        Close the workbook and quit the excel engine
        
        Parameters:
            None
            
        Returns:
            None
        '''

        self.engine.ScreenUpdating = True
        self.engine.Application.Quit()



    def save_and_quit(self) -> None:

        '''
        Save the workbook and quit the excel engine
        
        Parameters:
            None
            
        Returns:
            None
        
        Raises:
            FailedOnSaveException
        '''

        try:
            self.save()
            self.close_and_quit()
        except Exception as e:
            raise FailedOnSaveException()