import cv2
import numpy as np
import pyautogui
import random
import time
import xlrd
import win32clipboard
import pickle
messagetabx = 134
messagetaby = 31
saleswindowx = 28
saleswindowy = 19 #this is where i ended 1:30pm
checktimer = 0.2
def region_grabber (region):   
    x1 = region [0]
    y1 = region [1]
    width = region[2]-x1
    height = region[3] - y1
    
    return pyautogui.screenshot(region = (x1,y1,width,height))



def imagesearcharea(image, x1,y1,x2,y2, precision=0.8, im=None) :
    if im is None :
        im = region_grabber(region=(x1, y1, x2, y2))
        im.save('testarea.png') #usefull for debugging purposes, this will save the captured region as "testarea.png"

    img_rgb = np.array(im)
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    template = cv2.imread(image, 0)

    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val < precision:
        return [-1, -1]
    return max_loc

def click_image(image,x,y,  action, timestamp,offset=5):
    img = cv2.imread(image)
    #height, width, channels = img.shape
    #print(x)
    #print(y)
    #print (random.randint(-5,5))
    xcoord =x + (22 / 2)
    ycoord =y + (22 / 2)
    #print ("x coord: ",xcoord)
    #print ("y coord: ",ycoord)  
    pyautogui.moveTo(xcoord ,ycoord, timestamp)
    pyautogui.click(button=action)
    time.sleep(0.2)

#########################################

def OpenXlsx():
    ##### THINGS TO CHANGE #####
    print("OpenXlsx")
    #time.sleep(checktimer)
    book = xlrd.open_workbook('Edited October Sales.xls')
    sheet = book.sheet_by_name ('NANKID')
    things = [book, sheet]
    return things
    ######
    
def SaveNameData(name):
    #time.sleep(checktimer)
    print("SaveNameData")
    win32clipboard.OpenClipboard()
    name1 = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return name1

def SetNameData(name):
    #time.sleep(checktimer)
    print("SetNameData")
    while True:
            try:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(name)
                win32clipboard.CloseClipboard()

                name1 = SaveNameData(name)
                return name1
                break
            
            except pyautogui.FailSafeException as e:
                print ('failsafe!')
                break
        
            except Exception as error:
                print ("Clipboard problem. Restarting")
                time.sleep(0.2)

def ClickChatTab(messagetabx, messagetaby):
    #time.sleep(checktimer)
    print ("ClickChatTab")
    pyautogui.click(x = messagetabx, y = messagetaby)
    
def ClickSalesTab(saleswindowx, saleswindowy):
    #time.sleep(checktimer)
    print ("ClickSalesTab")
    pyautogui.click(x = saleswindowx, y = saleswindowy)

def SearchBoxClear(xsearchboxx, xsearchboxy):
    ClickSalesTab(saleswindowx, saleswindowy)
    time.sleep(checktimer)    
    print("SearchBoxClear")
    pyautogui.click (x = xsearchboxx, y = xsearchboxy)

def SearchBoxPaste(searchboxx, searchboxy):
    time.sleep(checktimer)
    print("SearchBoxPaste")
    #coordinates of search box
    pyautogui.click(x = searchboxx, y = searchboxy, clicks = 2)
    #paste on search box
    pyautogui.hotkey('ctrl','v')  
    #click inside search box

def CantFindChatIcon(image, x1, y1, x2, y2, precision, searchboxx, searchboxy, saleswindowx, saleswindowy):
    time.sleep(checktimer*2)
    print("CantFindChatIcon")
    ClickSalesTab(saleswindowx, saleswindowy)
    time.sleep(0.5)
    pyautogui.click(x = searchboxx, y = searchboxy)
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(0.3)
    pos = imagesearcharea(image, x1, y1, x2, y2, precision)

def CopyChatName():
    time.sleep(checktimer)
    print("CopyChatName")
    win32clipboard.OpenClipboard()
    name2 = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    name2 = name2.rstrip()
    return name2

def checkChatname(chatnamex, chatnamey):
    ClickChatTab(messagetabx, messagetaby)
    time.sleep(0.7)
    
    print("CheckChatName")
    pyautogui.click(x=chatnamex, y = chatnamey, clicks = 3)
    pyautogui.hotkey('ctrl','c')

    while True:
        try:
            name2 = CopyChatName()
            return name2
            break
        except Exception as error:
            print ("Clipboard problem. Restarting")
            time.sleep(0.2)

def CanFindChatIcon(image, searchareax1, searchareay1,pos):
    ClickSalesTab(saleswindowx, saleswindowy)
    time.sleep(checktimer)
    print("CanFindChatIcon")
    time.sleep(0.1)
    click_image(image,searchareax1 + pos[0], searchareay1+ pos[1],"left",0)
    time.sleep(0.5)

def FindChatIconTest(image, searchareax1, searchareay1,pos, x1, y1, x2, y2, precision, searchboxx, searchboxy,saleswindowx, saleswindowy):
    time.sleep(checktimer)
    ClickSalesTab(saleswindowx, saleswindowy)
    print("FindChatIcon")
    if pos[0] == -1: #cant find chat icon
        CantFindChatIcon(image, x1, y1, x2, y2, precision, searchboxx, searchboxy, saleswindowx, saleswindowy)                   
    elif pos[0] != -1: #can find chat icon
        CanFindChatIcon(image, searchareax1, searchareay1,pos)

def ClickFirstEntry(name, xsearchboxx, xsearchboxy, saleswindowx, saleswindowy, searchboxx, searchboxy, firstentryx, firstentryy):
    time.sleep(checktimer)
    print("ClickFirstEntry")
    ClickSalesTab(saleswindowx, saleswindowy)
    time.sleep(0.3)
    SearchBoxClear(xsearchboxx, xsearchboxy)
    SetNameData(name)
    SearchBoxPaste(searchboxx, searchboxy)
    time.sleep(0.5)
    pyautogui.click(x = firstentryx, y = firstentryy, clicks = 1)

def ClickChatBox(textboxx, textboxy):
    ClickChatTab(messagetabx, messagetaby)
    #time.sleep(checktimer)
    print("ClickChatBox")
    pyautogui.click(x = textboxx, y = textboxy, clicks = 2) #click text box

def SetChatData(ChatMessage):
    #time.sleep(checktimer)
    print("SetChatData")
    while True:
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(ChatMessage)
            win32clipboard.CloseClipboard()
            break
        
        except pyautogui.FailSafeException as e:
            print ('failsafe!')
            break
        
        except AssertionError as error:
            print ("Clipboard problem. Restarting")
            time.sleep(0.2)

def SendChatMessage(messagetabx, messagetaby):
    time.sleep(checktimer)
    #time.sleep(checktimer)
    pyautogui.click (x = messagetabx, y = messagetaby)
    print("SendChatMessage")
    pyautogui.hotkey('ctrl','v')
    #time.sleep(10)
    pyautogui.press('enter',presses = 2)   
    
def NamecounterErrorHolder(namecounter, namecounter2):
    time.sleep(checktimer)
    print ("NamecounterErrorHolder")
    if namecounter == None:
        print ("NamecounterErrorHolder: NoneType")
        namecounter = namecounter2 #set namecounter as the namecounter before (namecounter2)
        #namecounter2 = namecounter 
        namecounters = [namecounter, namecounter2]
        return namecounters
    else:
        print ("NamecounterErrorHolder: not NoneType")
        namecounter2 = namecounter #save namecounter 2 as the last namecounter
        namecounters = [namecounter, namecounter2]
        return namecounters

def ReadNamecounter():
    time.sleep(checktimer)
    print("ReadNameCounter")
    textstuff = "namecounter.txt"
    try:
        f = open (textstuff, "r+")
        namecounter = int(f.read())
        f.close()
        return namecounter
    except Exception as e:
        print (e)
        namecounter = 1
        print (namecounter)

def SaveNamecounter(namecounter):
    time.sleep(checktimer)
    print ("SaveNameCounter")
    Saveas = "namecounter.txt"
    f = open (Saveas, "w+")
    f.write (str(namecounter))
    f.close()

def ClearCache(cleanbuttonx, cleanbuttony, cachebuttonx, cachebuttony):
    pyautogui.click(x = cleanbuttonx, y =cleanbuttony)
    time.sleep(0.3)
    pyautogui.click(x = cachebuttonx, y = cachebuttony)
    time.sleep(1)
    
def FindRSearch(name1,n): #THIS WHOLE THING
    print("FindRSearch")
    ClickSalesTab(saleswindowx, saleswindowy)
    name = SetNameData(name1)
    pyautogui.hotkey('ctrl','shift','f')
    pyautogui.hotkey('backspace')
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    pyautogui.typewrite ('thing')
    pyautogui.click(522,409)
    pyautogui.click(651,324)
    time.sleep(1.5)
    

    
def ManualSearch(name1, n):
    print("ManualSearch")
    name = SetNameData(name1)
    pyautogui.hotkey('ctrl','f')
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter', presses = n)
    pyautogui.press('esc')
        

def ScrollTop():
    ClickSalesTab(saleswindowx, saleswindowy)
    pyautogui.click (193, 250)
    pyautogui.scroll(10000)

def ClickHighlightedTextChatButton(name): #THIS WHOLE THING
    print("ClickHighlightedTextChatButton")
    try:
        x, y = pyautogui.locateCenterOnScreen("CaptureThing.png", confidence=0.9)
        pyautogui.click(x+20,y)
        LogPassed(name+'\n')
    #except ImageNotFoundException as e:
    #    print ("Highlighted Text Chat Button Not Found")
    #    LogSkipped(name)
    #    exit
    except Exception as e:
        print ("Uncaught error...")
        print(e)
        LogSkipped(name+'\n')
        LogError(name+'>>'+str(e)+'\n')
        exit
    
def LogSkipped(name):
    print("LogSkipped")
    saveas = "skippedNames.txt"
    f = open(saveas, "a+")
    f.write(name)
    f.close()

def LogError(e):
    print("LogError")
    saveas = "errorLog.txt"
    f = open(saveas, "a+")
    f.write(e)
    f.close()
    
def LogPassed(name):
    print("LogPassed")
    saveas = "passedNames.txt"
    f = open(saveas, "a+")
    f.write(name)
    f.close()

##x1 = 700 
##y1 = 343
##x2 = 1346 
##y2 = 376
##    
##image = "Capture.PNG"
##pos = imagesearcharea(image, x1, y1, x2, y2, 0.8)
##CanFindChatIcon(image, x1, y1,pos)

##searchboxx = 784
##searchboxy = 239
##xsearchboxx = 943
##xsearchboxy = 241
##saleswindowx = 717
##saleswindowy = 29
##chatnamex = 938
##chatnamey = 116
##textboxx = 957
##textboxy = 649
##firstentryx = 824
##firstentryy = 386
##ClickFirstEntry("wheiame", xsearchboxx, xsearchboxy, saleswindowx, saleswindowy, searchboxx, searchboxy, firstentryx, firstentryy)
##
