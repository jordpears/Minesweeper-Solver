import time
import win32api, win32con
import cv2
import numpy as np
from PIL import ImageGrab, Image

#initial image importing for template matching and initial screenshot + saving
screen = ImageGrab.grab()
screen.save("images/screen.png")
screen = cv2.imread("images/screen.png")
topleft = cv2.imread("images/topleft.png")
bottomright = cv2.imread("images/bottomright.png")
    

def findgame(screen,template,n):
    #find the games' location on the screen.
    global bottomright,topleft
    templatemap = cv2.matchTemplate(screen,template,cv2.TM_CCOEFF)
    templatecoord = cv2.minMaxLoc(templatemap)
    templatecoord = templatecoord[3]
    if n == 0:
        templatecoord1 = (templatecoord[0] + 5,templatecoord[1] + 5)
    else:
        templatecoord1 = (templatecoord[0] + 21,templatecoord[1] + 21)    
    return templatecoord1
    
def gameboardsetup(tlcorner,brcorner,screen):
    #from screenshot of screen, crop to only be the screenshot of gameboard
    screen = Image.fromarray(screen, 'RGB')
    gameboard = screen.crop((tlcorner[0],tlcorner[1],brcorner[0],brcorner[1]))
    gameboard.load()
    return gameboard
    
    
def click(x,y):
    #left click the game
    xc = topleftcoord[0]+ 5 + (18 * x)
    yc = topleftcoord[1]+ 5 + (18 * y)
    win32api.SetCursorPos((xc,yc))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,xc,yc,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,xc,yc,0,0)
    time.sleep(0.2)
    win32api.SetCursorPos(topleftcoord)

def rightclick(x,y):
    #right click the game
    xc = topleftcoord[0]+ 5 + (18 * x)
    yc = topleftcoord[1]+ 5 + (18 * y)
    win32api.SetCursorPos((xc,yc))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,xc,yc,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,xc,yc,0,0)
    time.sleep(0.2)
    win32api.SetCursorPos(topleftcoord)
    
def findnumbers(x,y,gameboard):
    #identify numbers/unopened/flags on the gameboard screenshot
    b = 0 #correction for pixel devilry
    a = 0
    if x >= int(squares_x/2):
        b = 1
    if y >= int(squares_y/2):
        a = 1
    gameboard = np.asarray(gameboard)
    for number in range(-1,7):  
        if number == -1:
            if 33 < gameboard[(y*18)+17+a,(x*18)+9+b][2] < 60 and gameboard[(y*18)+15+a,(x*18)+11+b][2] < 80:
                return -1
        elif number == 0:
            if 180 > gameboard[(y*18)+15+a,(x*18)+13+b][1] > 165 :
                return -9
        elif number == 1:
            if np.array_equal(gameboard[(y*18)+11+a,(x*18)+10+b],(190,80,64)) and np.array_equal(gameboard[(y*18)+6+a,(x*18)+10+b],(190,80,64)) :
                return 1
        elif number == 2:
            if np.array_equal(gameboard[(y*18)+15+a,(x*18)+11+b],(0,104,31)):
                return 2
        elif number == 3:
            if 160 < gameboard[(y*18)+15+a,(x*18)+10+b][2] < 175:
                return 3
        elif number == 4:
            if np.array_equal(gameboard[(y*18)+10+a,(x*18)+12+b],(130,1,2)):
                return 4
        elif number == 5:
            if np.array_equal(gameboard[(y*18)+14+a,(x*18)+11+b],(3,5,121)):
                return 5
        else:
            return 0
        
        
def gameboardupdate():
    #update the gameboard image stored
    global topleftcoord,bottomrightcoord, gameboard
    screen = ImageGrab.grab()
    screen.save("images/screen.png")
    screen = cv2.imread("images/screen.png")
    screen = Image.fromarray(screen,"RGB")
    gameboard = screen.crop((topleftcoord[0],topleftcoord[1],bottomrightcoord[0],bottomrightcoord[1]))
    gameboard.load()
    return 

def strategy(gm):
    #logical deduction based on surrounding blocks.
    for y in range(0,len(gm)):
        for x in range(0,len(gm[0])):           
            countflag = 0
            countunopen = 0
            if gm[y][x] >= 1:
                for i in [-1,0,1]:
                    for j in [-1,0,1]:
                        try:
                            if gm[y+i][x+j] == -1 and y+i >= 0 and x+j >= 0:
                                countunopen += 1
                            if gm[y+i][x+j] == -9 and y+i >= 0 and x+j >= 0:
                                countflag += 1
                        except IndexError:
                            pass   
                
                if gm[y][x] == countflag:
                    for i in [-1,0,1]:
                        for j in [-1,0,1]:
                            try:
                                if gm[y+i][x+j] == -1 and y+i >= 0 and x+j >= 0:
                                    return (x+j),(y+i), False
                            except IndexError:
                                pass                  
                if gm[y][x] == countunopen + countflag:
                    for i in [-1,0,1]:
                        for j in [-1,0,1]:
                            try:
                                if gm[y+i][x+j] == -1 and y+i >= 0 and x+j >= 0:
                                    return (x+j),(y+i), True
                            except IndexError:
                                pass

    
#determine onscreen coords of game
topleftcoord = findgame(screen,topleft,0)  
bottomrightcoord = findgame(screen,bottomright,1)

#determine gameboard number of squares
gameboard = gameboardsetup(topleftcoord,bottomrightcoord,screen)
squares_x = int(gameboard.size[0]/18)
squares_y = int(gameboard.size[1]/18)

#setup game matrix
gamematrix = list(range(squares_y))
for i in range(squares_y):
    gamematrix[i] = list(range(squares_x))
for i in range(squares_y):
    for u in range(squares_x):        
        gamematrix[i][u] = -1 #game[y][x]
        
#gamefocus
click(0,0)

#firstclick
click(0,0) 
 
alive = True

#active loop
while alive == True:
    time.sleep(0.2)
    gameboardupdate() #refresh the gameboard matrix
    for y in range(squares_y):
        for x in range(squares_x):        
            gamematrix[y][x] = findnumbers(x,y,gameboard)#game[y][x]      
    x,y,flag = strategy(gamematrix)
    if flag == False:
        click(x,y)
    if flag == True:
        rightclick(x,y)
