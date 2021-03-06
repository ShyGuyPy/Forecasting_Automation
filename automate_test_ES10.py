
#test model is 'Markov Chain Weather' found in ExtendSim10/Exapmles/Continuous/Standard Block Models

# more information about setting up extend sim at:  https://docs.google.com/document/d/1u16zvcNphPmkxYZTrKp78__DTsLl4JpmrN62dfDhsXs/edit

import win32com.client
import win32gui as wg
import win32con
import time




import pyperclip
import clipboard
import pyautogui
PRRISM_url = "C:\\Users\icprbadmin\Desktop\PRRISM_v2p00_ops\PRRISM_v2.00_ops.mox"
cdm_url = "C:\\Users\icprbadmin\Desktop\PRRISM_v2p00_ops\library\CDM.lix"
newlibrary_url = "C:\\Users\icprbadmin\Desktop\PRRISM_v2p00_ops\library\\newlibrary.lix"

#--------------extend sim menu command codes required format----------------
#these, and more, can be found on  page 417 of ExendSimTechnicalReferenece.pdf
run_simulation = """ExecuteMenuCommand(6000)"""
stop_simulation = """ExecuteMenuCommand(30000)"""
pause_simulation = """ExecuteMenuCommand(30001)"""

recent_file_1 = """ExecuteMenuCommand(1555)"""
recent_file_2 = """ExecuteMenuCommand(1556)"""
recent_file_3 = """ExecuteMenuCommand(1557)"""
recent_file_4 = """ExecuteMenuCommand(1558)"""
recent_file_5 = """ExecuteMenuCommand(1559)"""

new_model =  """ExecuteMenuCommand(2)"""
open_model = """ExecuteMenuCommand(3)"""
close_model= """ExecuteMenuCommand(4)"""

save_model = """ExecuteMenuCommand(5)"""
save_model_as = """ExecuteMenuCommand(6)"""

properties = """ExecuteMenuCommand(2001)"""
exit_or_quit = """ExecuteMenuCommand(1)"""

cut = """ExecuteMenuCommand(18)"""
copy= """ExecuteMenuCommand(19)"""
paste = """ExecuteMenuCommand(20)"""
clear = """ExecuteMenuCommand(21)"""

#will make whichever version of extend sim you run this command in your default extend sim(will be the one run by this script)
update_launch_control = """ExecuteMenuCommand(1410)"""
#----------------------------------------------------------------------------

#use win32 to open Extend Sim and set window
def open_and_ID(prog_ID, win_ID):
    program_handle = win32com.client.Dispatch(prog_ID)
    app_ID = wg.FindWindow(None, win_ID)
    wg.ShowWindow(app_ID, win32con.SW_MAXIMIZE)
    wg.SetActiveWindow(app_ID)
    wg.SetForegroundWindow(app_ID)
    #print(program_handle)
    print(app_ID)
    return program_handle

#run currently open model(default model to open can be set in Extend Sim)
def run_by_id(prog_ID, win_ID):
    program_handle = win32com.client.Dispatch(prog_ID)
    app_ID = wg.FindWindow(None, win_ID)
    program_handle.Execute(run_simulation)#"""ExecuteMenuCommand(6000)""")

#set simulation parameters
def set_parameters(prog_ID, win_ID, SetEndTime, SetStartTime, SetNumSim, SetNumStep):
    program_handle = win32com.client.Dispatch(prog_ID)
    app_ID = wg.FindWindow(None, win_ID)
    # sets the setting parameters into a string that can be fed into the MODL execute
    execute_input = """SetRunParameters({}, {}, {}, {})""".format(SetEndTime, SetStartTime, SetNumSim, SetNumStep)
    program_handle.Execute(execute_input)

#this is just a wrapper for the sleep() function for clarity sake
def wait_time(x):
    time.sleep(x)

def test_click():
    print("click works")

#opens a file through the drop down menu
def open_file(prog_ID, win_ID):
    program_handle = win32com.client.Dispatch(prog_ID)
    program_handle.Execute(open_model)
    wait_time(3)
    pyperclip.copy(PRRISM_url)
    pyperclip.paste()

#open libraries, in sequence, required to open model
def open_libraries():

    clipboard.copy(cdm_url)
    #clipboard.paste()
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    # pyperclip.copy(cdm_url)
    # pyperclip.paste()
    # pyperclip.paste()
    # wait_time(10)
    # pyperclip.copy(newlibrary_url)
    # pyperclip.paste()

    wait_time(10)
    clipboard.copy(newlibrary_url)
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')


#open the most recently run model
def run_recent_1(prog_ID, win_ID):
    program_handle = win32com.client.Dispatch(prog_ID)
    app_ID = wg.FindWindow(None, win_ID)
    program_handle.Execute(recent_file_1 )#"""ExecuteMenuCommand(6000)""")


#open extend sim
es_handle = open_and_ID("Extend.application", "ExtendSim")

wait_time(2)

#opens a file through the drop down menu
#open_file("Extend.application", "ExtendSim")

#open the most recently run model
run_recent_1("Extend.application", "ExtendSim")

#wait_time(30)

#open_libraries()

wait_time(200)

#sets run parameters and then run the model
#function inputs are :prog_ID, win_ID, SetEndTime, SetStartTime, SetNumSim, SetNumStep
set_parameters("Extend.application", "ExtendSim", 1000, 0 , 1, 1) #30763

#run open model
run_by_id("Extend.application", "ExtendSim")