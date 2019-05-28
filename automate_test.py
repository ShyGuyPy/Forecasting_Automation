import win32com.client
import win32gui as wg
import win32con

def open_and_ID(prog_ID, win_ID):
    program_handle = win32com.client.Dispatch(prog_ID)
    app_ID = wg.FindWindow(None, win_ID)
    wg.ShowWindow(app_ID, win32con.SW_MAXIMIZE)
    wg.SetActiveWindow(app_ID)
    #print(program_handle)
    return program_handle


open_and_ID("Extend.application", "ExtendSim")

#com browser location
#C:\Users\coop_user\AppData\Local\Programs\Python\Python36-32\Lib\site-packages\win32com\client>python combrowse.py
