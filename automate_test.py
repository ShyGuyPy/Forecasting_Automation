import win32com.client
import win32gui as wg
import win32con
import time

def open_and_ID(prog_ID, win_ID):
    program_handle = win32com.client.Dispatch(prog_ID)
    app_ID = wg.FindWindow(None, win_ID)
    wg.ShowWindow(app_ID, win32con.SW_MAXIMIZE)
    wg.SetActiveWindow(app_ID)
    #print(program_handle)
    return program_handle

def wait_time(x):
    time.sleep(x)


es_handle = open_and_ID("Extend.application", "ExtendSim")
print(es_handle)
#wait_time(1)
#es_handle.GetObjectHandle("Extend.application", "ExtendSim")

#es_handle.GetObjectHandle("prototype1_v2.mox")#, 32) #, "lfalls_nat_9day")


es_handle.Execute("""ActivateApplication()""")
#brings specified worksheet to forefront
es_handle.Execute("""ActivateWorksheet("prototype1_v2.mox")""")
#set the run parameters SetEndTime, SetStartTime, SetNumSim, SetNumStep
es_handle.Execute(""" SetRunParameters(10000, 0 , 1, 1) """)
#runs the simulation, set to false to avaoid setup window pop up
es_handle.Execute(""" RunSimulation(false)""")


#test = es_handle.Request("test_model_1.mox", "E167B362-7044-11d2-99DE-00c0230406DF")
#print(test)
#es_handle(Poke(es_handle, "Processed_USGS.csv"))
#com browser location
#C:\Users\coop_user\AppData\Local\Programs\Python\Python36-32\Lib\site-packages\win32com\client>python combrowse.py
#under Registered Type Libraries
#under Extend Type Library
#check out Extend Application Dispatch
#check out all Extend Functions objects

#Extend Sim 8 GUID is: {E167B362-7044-11d2-99DE-00c0230406DF}

#model file prototype1_v2.mox can be found at H:\ICPRB2\USERS2\COOP\PRRISM\PRRISM_nextgen\prototype1