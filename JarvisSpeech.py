# Speech recognition

#Importing the necessary libraries
import speech_recognition as sr
import pyttsx3 as pts
import subprocess as sp
import os
import time
import winsound
import sys
import win32con
import win32api
import win32gui
import win32com.client
import pythoncom
from win32com.shell import shell, shellcon 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
 
# Getting the recognizer method and initialisation
recognizer = sr.Recognizer()
mic = sr.Microphone()
engine = pts.init()
frequency = 2500  # Set Frequency To 2500 Hertz
duration = 350  # Set duration to 350ms

#Function to open the actual file as per the user's choice
def openfile(number,file_paths):
    for data in range(len(file_paths)):
        if (data+1) == number :
            os.startfile(file_paths[data])
            
# Get list of paths from given Explorer window or from all Explorer windows.
def get_explorer_files( hwndOfExplorer = 0, selectedOnly = False ):
    print("Getting the list of files")
    paths = []

    # Create instance of IShellWindows (I couldn't find a constant in pywin32)
    CLSID_IShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
    shellwindows = win32com.client.Dispatch(CLSID_IShellWindows)

    # Loop over all currently open Explorer windows
    for window in shellwindows:
        # Skip windows we are not interested in.
        if hwndOfExplorer != 0 and hwndOfExplorer != window.HWnd:
            continue

        # Get IServiceProvider interface
        SP = window._oleobj_.QueryInterface( pythoncom.IID_IServiceProvider )

        # Query the IServiceProvider for IShellBrowser
        shBrowser = SP.QueryService( shell.SID_STopLevelBrowser, shell.IID_IShellBrowser )

        # Get the active IShellView object
        shView = shBrowser.QueryActiveShellView()
        
        # Get an IDataObject that contains the items of the view (either only selected or all). 
        aspect = shellcon.SVGIO_SELECTION if selectedOnly else shellcon.SVGIO_ALLVIEW
        items = shView.GetItemObject( aspect, pythoncom.IID_IDataObject )
        # Get the paths in drag-n-drop clipboard format. We don't actually use 
        # the clipboard, but this format makes it easy to extract the file paths.
        # Use CFSTR_SHELLIDLIST instead of CF_HDROP if you want to get ITEMIDLIST 
        # (aka PIDL) format, but you can't use the simple DragQueryFileW() API then. 
        data = items.GetData(( win32con.CF_HDROP, None, pythoncom.DVASPECT_CONTENT, -1, pythoncom.TYMED_HGLOBAL ))

        # Use drag-n-drop API to extract the individual paths.
        numPaths = shell.DragQueryFileW( data.data_handle, -1 )
        paths.extend([
                shell.DragQueryFileW( data.data_handle, i ) \
                for i in range( numPaths )
                ])

        if hwndOfExplorer != 0:
            break
    print("Sending the path of the file")
    engine.say("Which file do you wish to open ?")
    engine.runAndWait()
    engine.say("Mention the number of the file on the list, after the beep")
    engine.runAndWait()
    winsound.Beep(frequency, duration)
    recognizer.adjust_for_ambient_noise(source)
    audio = recognizer.listen(source,timeout=5,phrase_time_limit=6)
    engine.say("You just said :"+recognizer.recognize_google(audio))
    engine.runAndWait()
    openfile(int(recognizer.recognize_google(audio)),paths)
    
# Retrieving local data 
def retrieve(message):
    try:
        hwnd = 0  
        selectedOnly = False
        query_string = message
        local_path = r'C:\Users\Test'      
                                                 
        # For a local folder
        sp.Popen(f'explorer /root,"search-ms:query={query_string}&crumb=folder:{local_path}&"')    
        time.sleep(2)
        get_explorer_files(hwnd, selectedOnly)    
    
    except Exception as e:
            print( "ERROR: ", e )
            
#WEB ACCESS            
#Surfing the Web
def search_google(query):
    engine.say("You just said :"+query)
    engine.runAndWait()   
    browser = webdriver.Chrome(r"C:/Users/HP/Anaconda3/Scripts/chromedriver_win32/chromedriver.exe")
    browser.get('http://www.google.com')
    browser.maximize_window()
    search = browser.find_element_by_name('q')
    search.send_keys(query)
    search.send_keys(Keys.RETURN)
    time.sleep(6)
    engine.say("Say 'close' to close the browser at any time , after the beep.")
    engine.runAndWait()
    recognizer.adjust_for_ambient_noise(source)
    winsound.Beep(frequency, duration)
        
# Implementation  - Testing the catchphrase and file retrieval process
start = time.time()
with mic as source:
    recognizer.adjust_for_ambient_noise(source)
    print("Say 'hello jarvis' to begin recording audio")
    audio = recognizer.listen(source,timeout=5,phrase_time_limit=3)
    word = 'hello Jarvis'
    if word in recognizer.recognize_google(audio):
            engine.say("Hello , I am JARVIS , your voice assistant")
            engine.runAndWait()
            engine.say("Say 'Google' to access the internet and , 'PC' to access local data ,  after the beep.")
            engine.runAndWait()
            net = 'Google'
            file = 'PC'
            winsound.Beep(frequency, duration)
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source,timeout=5,phrase_time_limit=2)
            if net in recognizer.recognize_google(audio):
                engine.say("Mention what should be searched for , after the beep")
                engine.runAndWait()
                winsound.Beep(frequency, duration)
                recognizer.adjust_for_ambient_noise(source)
                audio = recognizer.listen(source,timeout=5,phrase_time_limit=6)
                search_google(recognizer.recognize_google(audio))                 
            elif file in recognizer.recognize_google(audio):
                engine.say("Mention the name of the file to be accessed , after the beep")
                engine.runAndWait()
                winsound.Beep(frequency, duration)
                recognizer.adjust_for_ambient_noise(source)
                audio = recognizer.listen(source,timeout=4,phrase_time_limit=5)
                retrieve(recognizer.recognize_google(audio))
end = time.time()
print("Execution time:",(end-start))            









