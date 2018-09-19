import ctypes 
import win32gui
import win32con
import win32api
import win32process 
import win32com.client
import time

#引入winapi
gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
shell = win32com.client.Dispatch("WScript.Shell")

def getHandle():
    return win32gui.FindWindow("Minesweeper", None)

def getWindowRect(handle):
    return win32gui.GetWindowRect(handle)

def openProcess(handle):
    PROCESS_ALL_ACCESS=(0x000F0000|0x00100000|0xFFF) 
    pid=win32process.GetWindowThreadProcessId(handle)[1] #根据窗体抓取进程编号
    return win32api.OpenProcess(PROCESS_ALL_ACCESS,False,pid)#用最高权限打开进程编号

def click(rect,i,j):
    x = rect[0]+15+j*16+8
    y = rect[1]+100+i*16+8
    win32api.SetCursorPos([x,y])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP  | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

def isSafe():
    hdc = user32.GetDC(None)
    c = gdi32.GetPixel(hdc,0,0)
    return c==16777215

def isSafeEx(phand,i,j):
    date=ctypes.c_ubyte()
    address = 0x1005361+0x20*i+j
    kernel32.ReadProcessMemory(int(phand), address, ctypes.byref(date), 1, None)
    return date.value

def getMinesNum(phand):
    date=ctypes.c_ubyte()
    kernel32.ReadProcessMemory(int(phand), 0x1005330, ctypes.byref(date), 1, None)
    return int(date.value)

def getMapSize(phand):
    height=ctypes.c_byte()
    width=ctypes.c_byte()
    kernel32.ReadProcessMemory(int(phand), 0x1005338, ctypes.byref(height), 1, None)
    kernel32.ReadProcessMemory(int(phand), 0x1005334, ctypes.byref(width), 1, None)
    return [int(width.value),int(height.value)]

def calGameDate(gameDate,safeBlocks,mapSize):
    pos = [ [0,-1],[1,-1],[1,0],[1,1],[0,1],[-1,1],[-1,0],[-1,-1]]
    for i,j in safeBlocks:
        for pi,pj in pos:
            ni,nj = pi+i,pj+j
            if(ni<0 or nj<0 or ni>=mapSize[1] or nj >= mapSize[0] or gameDate[ni][nj]!=-1):
                continue
            gameDate[i][j]+=1

handle = getHandle() #游戏窗口句柄
if(handle == 0): 
    print("Minesweeper Not Found!")
    exit
while True:
    input("Press any key to continue...")
    handle = getHandle()
    if(handle == 0):
        print("Minesweeper Not Found!")
        continue
    phand = openProcess(handle) #打开游戏进程
    win32gui.ShowWindow(handle, win32con.SW_SHOWNORMAL)#激活
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(handle)#获取焦点
    winRect = getWindowRect(handle) #获取窗口位置
    mapSize = getMapSize(phand)
    print("GameWindRect:" + str(winRect))
    print("MapSize:" + str(mapSize) + " Mines:" + str(getMinesNum(phand)))
    safeBlocks = []
    mines = []
    gameDate = []
    for i in range(0,mapSize[1]):
        tmp = []
        for j in range(0,mapSize[0]):
            flag = isSafeEx(phand,i,j)
            if(flag == 0xf):
                tmp.append(0)
                safeBlocks.append([i,j])
            else:
                tmp.append(-1)
                mines.append([i,j])
        gameDate.append(tmp)
    
    calGameDate(gameDate,safeBlocks,mapSize)

    for i in range(0,mapSize[1]):
        for j in range(0,mapSize[0]):
            if(gameDate[i][j] == -1):
                print(' *',end='')
            else:
                print('%2d'%gameDate[i][j],end='')
        print('')
                
    for i,j in safeBlocks:
        click(winRect,i,j)
    print("Complete...")

    

