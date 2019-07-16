import win32com.client
import win32gui as wg
import win32con
import time

def open_and_ID(prog_ID, win_ID):
    program_handle = win32com.client.Dispatch(prog_ID)
    app_ID = wg.FindWindow(None, win_ID)
    wg.ShowWindow(app_ID, win32con.SW_MAXIMIZE)
    wg.SetActiveWindow(app_ID)
    wg.SetForegroundWindow(app_ID)
    #print(program_handle)
    print(app_ID)
    return program_handle

def wait_time(x):
    time.sleep(x)


es_handle = open_and_ID("Extend.application", "ExtendSim")

#set the run parameters SetEndTime, SetStartTime, SetNumSim, SetNumStep
es_handle.Execute(""" SetRunParameters(10000, 0 , 1, 1) """)

