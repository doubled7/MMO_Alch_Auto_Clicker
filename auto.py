from json import dump, loads
from os import getenv, makedirs, path
from random import uniform as random_between
from tkinter import Button, Entry, Label, StringVar, Tk

from keyboard import on_press
from win32api import (GetSystemMetrics,  # pylint:disable=no-name-in-module
                      SetCursorPos, mouse_event)
from win32con import MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
from win32gui import GetCursorInfo  # pylint:disable=no-name-in-module

# Options
FOCUS_WINDOW_ON_TOGGLE = False
UPDATE_DELAY_MILLISECONDS = 200 # UPDATE delay for current x/y coordinates in GUI
WINDOW_TITLE = 'MMO Auto Clicker'

# User Info
settingsDir = getenv('APPDATA') +'/AutoClickerSettings'
settingsFilePath = settingsDir +'/settings.json'
# RESOLUTION_WIDTH = GetSystemMetrics(0)
# RESOLUTION_HEIGHT = GetSystemMetrics(1)

def clampTimeLowerBound(n): 
    ''' Make sure the delay is at least 1 MS '''
    return max( 1, n )

def clamp(n, smallest, largest): return max(smallest, min(n, largest))

def tryGetAdjInt(formControl):
    ''' Checks whether user input is valid '''

    try:
        value = int(formControl.get())
        adjTime = clampTimeLowerBound(value)
    except Exception as e:
        # print("Your user input was invalid!")
        return False
    
    return adjTime


class AutoClicker:
    ''' Creates an autoclicker '''

    AUTO_ACTIVE = False
    START_KEY_BIND_OPEN = False     # Whether or not we are recording the next keystroke to update binds
    COORDS_KEY_BIND_OPEN = False

    settings = {
        'MIN_DELAY': 1400,
        'MAX_DELAY': 1600 ,
        'CURSOR_RANDOM_X_POSITION_WIDTH': 3,          # how many pixels +- should we move from the set cursor position
        'CURSOR_RANDOM_Y_POSITION_HEIGHT': 1,
        'COORDS_X': 0,
        'COORDS_Y': 0,
        'KEY_TOGGLE': 'page up',
        'KEY_MARK': 'end',
    }


    def updateGUI(self):
        opt = self.settings

        self.minTextValue.set(opt['MIN_DELAY'])
        self.maxTextValue.set(opt['MAX_DELAY'])
        self.widthTextValue.set(opt['CURSOR_RANDOM_X_POSITION_WIDTH'])
        self.heightTextValue.set(opt['CURSOR_RANDOM_Y_POSITION_HEIGHT'])
        self.xTextValue.set(opt['COORDS_X'])
        self.yTextValue.set(opt['COORDS_Y'])
        self.startKeyValue.set(opt['KEY_TOGGLE'])
        self.markKeyValue.set(opt['KEY_MARK'])
    
    def loadSettings(self):
        if not path.exists(settingsDir):
            makedirs(settingsDir)

        if not path.isfile(settingsFilePath):       # no user set options 
            self.updateGUI()
            return
        
        try:
            with open(settingsFilePath, 'r', encoding='utf-8') as f:
                self.settings = loads(f.read())
                f.close()
        except IOError:
            print("The settings file could not be read!")

        self.updateGUI()


    def saveSettingsOnClose(self):
        with open(settingsFilePath, 'w') as f:
            dump(self.settings, f)
            f.close()

        self.window.destroy()

    def toggleActiveStatus(self):
        if not self.AUTO_ACTIVE:
            self.btn.configure(text="Stop")
            self.AUTO_ACTIVE = True
            self.window.wm_title(WINDOW_TITLE +' *active*')
        else:
            self.btn.configure(text="Start")
            self.AUTO_ACTIVE = False
            self.window.wm_title(WINDOW_TITLE)

        # Update for both buttons
        self.attemptUpdateValues()


    def handleKeyPress(self, key):
        if key.name == self.settings['KEY_TOGGLE']:
            # TODO self.window.focus_force()
            self.toggleActiveStatus()
        elif key.name == self.settings['KEY_MARK']:
            self.markCoordinates()
        elif key.name == 'enter':
            self.attemptUpdateValues()
        
        if self.START_KEY_BIND_OPEN:
            self.settings['KEY_TOGGLE'] = key.name
            self.START_KEY_BIND_OPEN = False
            self.updateGUI() 
        elif self.COORDS_KEY_BIND_OPEN:
            self.settings['KEY_MARK'] = key.name
            self.COORDS_KEY_BIND_OPEN = False
            self.updateGUI()

    def handleToggleBindPress(self):
        self.COORDS_KEY_BIND_OPEN = False
        self.START_KEY_BIND_OPEN = True
    
    def handleMarkBindPress(self):
        self.START_KEY_BIND_OPEN = False
        self.COORDS_KEY_BIND_OPEN = True

    def execMouseClick(self):
        if self.AUTO_ACTIVE:
            width = self.settings['CURSOR_RANDOM_X_POSITION_WIDTH']
            height = self.settings['CURSOR_RANDOM_Y_POSITION_HEIGHT']
            x = self.settings['COORDS_X'] +int(random_between(width*-1, width))
            y = self.settings['COORDS_Y'] +int(random_between(height*-1, height))
            SetCursorPos((x,y))
            mouse_event(MOUSEEVENTF_LEFTDOWN,x,y,0,0)
            mouse_event(MOUSEEVENTF_LEFTUP,x,y,0,0)

        minTime = clampTimeLowerBound(self.settings['MIN_DELAY']) 
        maxTime = clampTimeLowerBound(self.settings['MAX_DELAY'])
        nextClickTime = int(random_between(minTime, maxTime))

        # print ("CLick: %s " %(str(nextClickTime)))

        self.window.after(nextClickTime, self.execMouseClick)
    # execMouseClick(50,10)

    def markCoordinates(self):
        ''' Updates the X,Y coordinates used '''

        flags, hcursor, (x,y) = GetCursorInfo()
        self.settings['COORDS_X'] = x
        self.settings['COORDS_Y'] = y
        self.updateGUI()

    def attemptUpdateValues(self):
        ''' Tries to convert user input into usable numbers if it's valid '''

        fMin = tryGetAdjInt(self.minTextValue)
        fMax = tryGetAdjInt(self.maxTextInput)
        if not fMin or not fMax or not fMin < fMax:
            errors = "The min/max is incorrect!"
            self.errorLabel.configure(text=errors)
            return 
        
        try:
            intWidthRandom = max( 0, int(self.widthTextInput.get()) )
            intHeightRandom = max( 0, int(self.heightTextInput.get()) )
            intXStart = int(self.xTextInput.get())
            intYStart = int(self.yTextInput.get())
        except Exception as e:
            errors = "One of more of your inputs was incorrect!"
            self.errorLabel.configure(text=errors)
            return False

        key_toggle = self.settings['KEY_TOGGLE']
        key_mark = self.settings['KEY_MARK']
        self.settings = {
            'MIN_DELAY': fMin,
            'MAX_DELAY': fMax,
            'CURSOR_RANDOM_X_POSITION_WIDTH': intWidthRandom,          # how many pixels +- should we move from the set cursor position
            'CURSOR_RANDOM_Y_POSITION_HEIGHT': intHeightRandom,
            'COORDS_X': intXStart,
            'COORDS_Y': intYStart,
            'KEY_TOGGLE': key_toggle,
            'KEY_MARK': key_mark,
        }
        self.errorLabel.configure(text="") # Remove previous errors if valid input
        self.updateGUI()

    def updateTempCoordinates(self):
        ''' This updates the X and Y coordinates that change when you move your cursor '''
        flags, hcursor, (x,y) = GetCursorInfo()
        self.curXLabel.configure(text=str(x))
        self.curYLabel.configure(text=str(y))

        self.window.after(UPDATE_DELAY_MILLISECONDS, self.updateTempCoordinates)

    def loadWindowPipeline(self):
        # Create the window
        self.window = Tk()
        # self.window.attributes("-topmost", True)
        self.window.lift()
        self.window.title(WINDOW_TITLE)
        self.window.geometry('410x180')

        # Add Controls/Labels
        self.curXDesc = Label(self.window, text="Current X: " )
        self.curXDesc.grid(column=0, row=0)
        self.curXLabel = Label(self.window, text="" )
        self.curXLabel.grid(column=1, row=0)
        
        self.curYDesc = Label(self.window, text="Current Y: " )
        self.curYDesc.grid(column=2, row=0)
        self.curYLabel = Label(self.window, text="" )
        self.curYLabel.grid(column=3, row=0)
        #####################################################################


        self.minLabel = Label(self.window, text="Min Delay: " )
        self.minLabel.grid(column=0, row=1)
        self.minTextValue = StringVar()
        self.minTextInput = Entry(self.window, width=14, textvariable=self.minTextValue)
        self.minTextInput.grid(column=1, row=1)
        

        self.maxLabel = Label(self.window, text="Max Delay: " )
        self.maxLabel.grid(column=2, row=1)
        self.maxTextValue = StringVar()
        self.maxTextInput = Entry(self.window, width=14, textvariable=self.maxTextValue)
        self.maxTextInput.grid(column=3, row=1)
        
        ####################################################################
        self.widthLabel = Label(self.window, text="Width: " )
        self.widthLabel.grid(column=0, row=2)
        self.widthTextValue = StringVar()
        self.widthTextInput = Entry(self.window, width=14, textvariable=self.widthTextValue)
        self.widthTextInput.grid(column=1, row=2)
        

        self.heightLabel = Label(self.window, text="Height: " )
        self.heightLabel.grid(column=2, row=2)
        self.heightTextValue = StringVar()
        self.heightTextInput = Entry(self.window, width=14, textvariable=self.heightTextValue)
        self.heightTextInput.grid(column=3, row=2)
        ####################################################################
        self.xLabel = Label(self.window, text="Start X: " )
        self.xLabel.grid(column=0, row=3)
        self.xTextValue = StringVar()
        self.xTextInput = Entry(self.window, width=14, textvariable=self.xTextValue)
        self.xTextInput.grid(column=1, row=3)
        

        self.yLabel = Label(self.window, text="Start Y: " )
        self.yLabel.grid(column=2, row=3)
        self.yTextValue = StringVar()
        self.yTextInput = Entry(self.window, width=14,textvariable=self.yTextValue)
        self.yTextInput.grid(column=3, row=3)
        
        ####################################################################
        self.startKeyLabel = Label(self.window, text="Toggle Key: " )
        self.startKeyLabel.grid(column=0, row=4)
        self.startKeyValue = StringVar()
        self.startKeyInput = Button(self.window, width=11,command=self.handleToggleBindPress, textvariable=self.startKeyValue)
        self.startKeyInput.grid(column=1, row=4)


        self.updateBtn = Button(self.window, text="Update", command=self.attemptUpdateValues, width=11)
        self.updateBtn.grid(column=3, row=4)
        ####################################################################
        self.markKeyLabel = Label(self.window, text="Mark X,Y Key: " )
        self.markKeyLabel.grid(column=0, row=5)
        self.markKeyValue = StringVar()
        self.markKeyInput = Button(self.window, width=11, command=self.handleMarkBindPress, textvariable=self.markKeyValue)
        self.markKeyInput.grid(column=1, row=5)
        ####################################################################
        self.errorLabel = Label(self.window, text="", fg="Red")
        self.errorLabel.place(x=2, y=148)
        ####################################################################
        # self.descLabel1 = Label(self.window, text="Use 'end' to mark coordinates")
        # self.descLabel1.place(x=20, y=106)
        # ####################################################################
        # self.descLabel2 = Label(self.window, text="Use 'page up' to toggle")
        # self.descLabel2.place(x=20, y=126)

        self.btnLabel = Label(self.window, text="Turn on/off" )
        self.btnLabel.grid(column=2, row=6)
        self.btn = Button(self.window, text="Start", command=self.toggleActiveStatus, width=11)
        self.btn.grid(column=3, row=6)

        ## Configure Grid
        self.window.grid_columnconfigure(0, minsize=70) 
        self.window.grid_columnconfigure(1, minsize=110) 
        self.window.grid_columnconfigure(2, minsize=110) 

        # Finish Loading
        self.loadSettings()
        self.updateTempCoordinates()
        self.execMouseClick()
        
        return self.window


## Initialization
ac = AutoClicker()
window = ac.loadWindowPipeline()

## Events
on_press(ac.handleKeyPress)
window.protocol("WM_DELETE_WINDOW", ac.saveSettingsOnClose)

## Main loop
window.mainloop()
