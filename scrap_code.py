 ###tomorrow...
 # run C:\Users\icprbadmin\AppData\Local\Continuum\anaconda3\Lib\site-packages\win32com\client>python combrowse.py in
 #terminal to access combrowse.py interface and find ExtednSim10 explicit idispatch #
 #  it's IID:{E167B361-7044-11D2-99DE-00C0230406DF} for our extendsim10
 #also seems to have built in cursor and wheel control (?)
#app id is {"6D809377-6AF0-444B-8957-A3773F02200E"}

#calls function(above) that opens Extend SIm and maximizes window


#"6D809377-6AF0-444B-8957-A3773F02200E")

#, "ExtendSim10")


#brings Extend Sim to the front
# es_handle.Execute("""ActivateApplication()""")


#wait_time(5)

#test = es_handle.Execute("""FileExists("prototype1_v2.mox")""")

#test = es_handle.Execute("""FileOpen("C:\\Users\icprbadmin\Documents\ExtendSim10\Examples\Continuous\Standard Block Models\test1.mox","")""")


#brings specified worksheet to forefront
#test2 = es_handle.Execute("""ActivateWorksheet("test_model.mox")""")
#print(test2)



#runs the simulation, set to false to avaoid setup window pop up
###es_handle.Execute(""" RunSimulation(false)""")

#quick test to make sure app_id was being assigned correctly...it is.  problem likely lies somewhere adjacent to this
#run tet again tomorrow for refresher
#also look at https://stackoverflow.com/questions/6312627/windows-7-how-to-bring-a-window-to-the-front-no-matter-what-other-window-has-fo
#aslo review win32gui package

# wait_time(5)
#
# fgwin = wg.GetForegroundWindow()
# print("app_id is: {}".format(fgwin))




#wait_time(1)
#test = es_handle.GetObjectHandle("Extend.application", "ExtendSim")
#print(test)

#es_handle.GetObjectHandle("prototype1_v2.mox")#, 32) #, "lfalls_nat_9day")
#es_handle.Execute("""FileOpen("Users\icprbadmin\Documents\Python_Scripts\Forecasting_Automation\prototype1_v2.mox", "Idunno")""")

###need to open the prototype1_v2.mox model here

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