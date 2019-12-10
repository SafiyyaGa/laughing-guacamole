import pyautogui
import FunctionCode
import Parts
import traceback
namecounter = 1
global searchareax1
searchareax1 = 238
global searchareay1 
searchareay1 = 423
global searchareax2 
searchareax2 = 820
global searchareay2 
searchareay2 = 459
cleanbuttonx = 1267
cleanbuttony =64
cachebuttonx =1112
cachebuttony =206

temperrornamecounter = 0


things = Parts.OpenXlsx()
print (things[1])

sheet = things[1]
book = things [0]
#namecounter = 12079
data = [sheet.cell_value(r, 0) for r in range (0, sheet.nrows)]#(sheet.nrows)]
namecounter2 = 0
ErrorCounter = 0
while namecounter !=sheet.nrows:
    try:
        if namecounter%50== 0:  
            print("is a multiplier of 5")
            #Parts.ClearCache(cleanbuttonx, cleanbuttony, cachebuttonx, cachebuttony)
            namecounter = FunctionCode.imagesearch_region_loop(things, \
                                                               namecounter,\
                                                               "Capture.PNG",\
                                                               '0.1',\
                                                               searchareax1, \
                                                               searchareay1, \
                                                               searchareax2, \
                                                               searchareay2)
            namecounters = Parts.NamecounterErrorHolder(namecounter, namecounter2)
            namecounter2 = namecounters[1]
            namecounter = namecounters[0]
            
            raise Exception
        
        elif namecounter <= sheet.nrows:
            print ('hello')
            namecounter = FunctionCode.imagesearch_region_loop(things,\
                                                               namecounter,\
                                                               "Capture.PNG",\
                                                               '0.1',\
                                                               searchareax1, \
                                                               searchareay1, \
                                                               searchareax2, \
                                                               searchareay2)
            namecounters = Parts.NamecounterErrorHolder(namecounter, namecounter2)

            namecounter2 = namecounters[1]
            namecounter = namecounters[0]
            print ("Error Counter = " & ErrorCounter) #THIS
            print ("TempNameCounter = " & temperrornamecounter)#THIS

            print ("is this you?",namecounter)
            #namecounter = namecounter+1        

        
        #elif namecounter >10:
            #pyautogui.moveTo(0,0)
            
    except pyautogui.FailSafeException as e:
        print ('Failsafe!')
        break
    
    except Exception as a:
        if temperrornamecounter == namecounter:#THIS
            if ErrorCounter < 10:#THIS
                print ("ErrorCounter", ErrorCounter)
                ErrorCounter = ErrorCounter + 1
            
                print (a)
                print (traceback.format_exc())
                print ('exception')
                namecounter = namecounter2
            else:
                print ("Code is movin on due to 10 consecutive Errors") #THIS
                namecounter = namecounter2 + 1#THIS
        else:#THIS
            temperrornamecounter = namecounter#THIS
            ErrorCounter = 0#THIS
        
        #continue
        #pass




