import win32com.client as win32

# ensure a new excel engine
excel_engine = win32.gencache.EnsureDispatch("Excel.Application")