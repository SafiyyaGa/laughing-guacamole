import cv2
import numpy as np
import pyautogui
import random
import time
import xlrd
import win32clipboard
import pickle
import Parts

###############

    
def imagesearch_region_loop(things, \
                            namecounter, \
                            image, \
                            timesample,\
                            x1, \
                            y1,\
                            x2, \
                            y2, \
                            precision=0.8):
    
    #RepeatTimes = 3
    #i = 0
    
    sheet = things[1]
    book = things [0]
    namecounter = Parts.ReadNamecounter()
    data = [sheet.cell_value(r, 0) for r in range (0, sheet.nrows)]#(sheet.nrows)]


#####THINGS TO CHANGE #####
    ChatMessage = "It's 11.11 and Shopee and Nankid have a treat for you! Get \
10% off on Nankid products when you use the code \
SHOPEEBABY1111 with a minimum spend of Php 500! \
\n\nThis promo valid only today, November 11, 2019!"		
		

    searchboxx = 246
    searchboxy = 240
    xsearchboxx = 450
    xsearchboxy = 247
    saleswindowx = 28
    saleswindowy = 19 #this is where i ended 1:30pm
    chatnamex = 311
    chatnamey = 157
    textboxx = 319
    textboxy = 627
##    firstentryx = 824
##    firstentryy = 386
    messagetabx = 134
    messagetaby = 31
    global searchareax1
    searchareax1 = x1
    global searchareay1 
    searchareay1 = y1
    global searchareax2 
    searchareax2 = x2
    global searchareay2 
    searchareay2 = y2
    #sheet.nrows
#########################################
    
    #while namecounter != sheet.nrows:
        
    name1 = Parts.SetNameData(data[namecounter]) #sets name into clipboard
    print(name1)
    Parts.ScrollTop()
    Parts.SearchBoxClear(xsearchboxx, xsearchboxy) #clears search box
    Parts.SearchBoxPaste(searchboxx, searchboxy) #clicks search box
    pyautogui.press('enter',presses = 1)
    time.sleep(1.5)
    pyautogui.click(x = 193, y = 250,clicks = 2)
    time.sleep(1)
    pyautogui.press('enter',presses=1)
    time.sleep(1)
    #pyautogui.click(x = searchboxx, y = searchboxy)
    #time.sleep(0.1)
    #search
    
    
    pos = Parts.imagesearcharea(image, x1,y1,x2,y2, precision)
    if pos[0] == -1:
        Parts.CantFindChatIcon(image, x1, y1, x2, y2, precision, searchboxx, searchboxy, saleswindowx, saleswindowy)
    elif pos[0] != -1: #else if
        #clicks image
        Parts.click_image(image,searchareax1 + pos[0], searchareay1+ pos[1],"left",0)
        time.sleep(1)
        print ('it got to here')
        #click on chat name

        #time.sleep(4) #checking timer
        name2 = Parts.checkChatname(chatnamex, chatnamey)
        n = 1
        while name2 != name1:
            print("Name1: ",name1," with len ",len(name1), " is not equal to")
            print("name2 ",name2," with len ",len(name2))
            name2 = Parts.checkChatname(chatnamex, chatnamey)
            #click on chat name
            #print (len(name1), len (name2))
            print("n: ", n)
            n = n + 1 #counter for repeats

            if name2 == name1:
                #print (len(name1), len (name2))
                print ('Name in DB is now equal to name in Chat Tab')
                break

            if name2 != name1: 
                if n <= 3:
                    print ("Finding Chat Icon Test (< 3)")
                    pyautogui.click(x = saleswindowx, y = saleswindowy)
                    pos = Parts.imagesearcharea(image, x1, y1, x2, y2, precision)
                    Parts.FindChatIconTest (image, searchareax1, searchareay1,pos, x1, y1, x2, y2, precision, searchboxx, searchboxy, saleswindowx, saleswindowy)
                    continue 
            
                elif n >3 :
                    print ("Finding Chat Icon Test (> 3)")
                    #n = 0
                    #Parts.ClickFirstEntry(data[namecounter], xsearchboxx, xsearchboxy, saleswindowx, saleswindowy, searchboxx, searchboxy, firstentryx, firstentryy)
                    #time.sleep(1.5)
                    #pos = Parts.imagesearcharea(image, x1, y1, x2, y2, precision)
                    #Parts.FindChatIconTest(image, searchareax1, searchareay1, pos, x1, y1, x2, y2, precision, searchboxx, searchboxy, saleswindowx, saleswindowy)
                    pyautogui.click(x = saleswindowx, y = saleswindowy)
                    Parts.ScrollTop()
                    # NEW 11/11/2019 3:35AM
                    pyautogui.click(22,716) # scroll left
                    pyautogui.click(232,181) # All button
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(name1)
                    win32clipboard.CloseClipboard()
                    Parts.SearchBoxClear(xsearchboxx, xsearchboxy) #clears search box
                    Parts.SearchBoxPaste(searchboxx, searchboxy) #clicks search box
                    time.sleep(1)
                    pyautogui.press('enter',presses = 1)
                    time.sleep (1)
                    ####
                    
                    i = 1
                    #Parts.ManualSearch(name1, i)
                    Parts.FindRSearch(name1,i)
                    Parts.ClickHighlightedTextChatButton(name1)
                    time.sleep(1)#THIS
                    i = i + 1
                    if i > 7:
                        Parts.LogSkipped(name1 + '\n')
                        break
                    #if i > 2:
                    #    pyautogui.scroll(-5)
                    continue
                elif n>5:
                    print("Finding Chat Icon Test (>5)")
                    print("Discontinuing sequence")
                    break
                
        time.sleep(0.5)
        Parts.ClickChatTab(messagetabx, messagetaby)
        Parts.ClickChatBox(textboxx, textboxy)
        Parts.SetChatData(ChatMessage)
        Parts.SetChatData(ChatMessage)
        Parts.SendChatMessage(messagetabx, messagetaby)
        
        #clicks back to sales
        pyautogui.click(x=saleswindowx, y =saleswindowy, clicks = 1)                
        print("NAMECOUNTER,", namecounter)
        namecounter = namecounter + 1
        print ("NEXT NAMECOUNTER", namecounter)
        Parts.SaveNamecounter(namecounter)
        return namecounter

### things to change ##########
##global searchareax1
##searchareax1 = 700
##global searchareay1 
##searchareay1 = 343
##global searchareax2 
##searchareax2 = 1346
##global searchareay2 
##searchareay2 = 376

####################

#imagesearch_region_loop(namecounter,"Capture.PNG",'0.1',searchareax1, searchareay1, searchareax2, searchareay2)
