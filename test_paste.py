import clipboard
import pyautogui
import time

cdm_url = "C:\\Users\icprbadmin\Desktop\PRRISM_v2p00_ops\library\CDM.lix"
newlibrary_url = "C:\\Users\icprbadmin\Desktop\PRRISM_v2p00_ops\library\\newlibrary.lix"


def wait_time(x):
    time.sleep(x)

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

    # wait_time(10)
    # clipboard.copy(newlibrary_url)
    # pyautogui.hotkey('ctrl','v')
    # pyautogui.press('enter')


wait_time(5)

open_libraries()
