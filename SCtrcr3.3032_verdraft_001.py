import openpyxl
#from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook

import tkinter as tk
from tkinter import ttk

import os, sys,shutil

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService

root = tk.Tk()
root.title("SoundCloud track crawler")
root.geometry('890x330')
canvas1 = tk.Canvas(root, width=800, height=730)
canvas1.pack(padx=3,pady=3)

ListOfEntries=[]

fixed_entry = ttk.Entry(canvas1, width=100)
fixed_entry.pack(padx=3,pady=3)
fixed_entry.focus()

text_entryNumbers=[]

def addurlbox():
    global extraentryBoxValue
    global ListOfAddedEntryBoxes
    ListOfAddedEntryBoxes=[]
    extraentryBox = tk.Entry(canvas1,width=100)
    extraentryBox.pack(padx=3,pady=3)
    text_entryNumbers.append(extraentryBox)
    extraentryBoxValue=extraentryBox
    ListOfAddedEntryBoxes.append(extraentryBox)

def crawler():
    inputs = []
    for widget in canvas1.winfo_children():
        if isinstance(widget, tk.Entry):
            inputs.append(widget.get())

    for i in range(len(inputs)):
        # taking url of SoundCloud platform
        # user inputs it, as string value
        varUrl_SC = inputs[i]

        # going to url(SC) via Selenium WebDriver
        chrome_options = Options()
        chrome_options.headless = True
        chrome_options.add_argument("start-maximized")
        # options.add_experimental_option("detach", True)
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')

        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.dirname(__file__)
            return os.path.join(base_path, relative_path)

        #driver = webdriver.Chrome(resource_path('./driver/chromedriver.exe'))

        #webdriver_service = Service(resource_path('./driver/chromedriver.exe'))
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        driver.get(varUrl_SC)

        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))).click()
            print('accepted cookies')
        except Exception as e:
            print('no cookie button!')

        # get text value of below items from opened SC website:
        # soundTitle_username, title, duration
        soundTitle_usernameTitleContainer = driver.find_element(By.CLASS_NAME, "soundTitle__title")
        varStrTitle = soundTitle_usernameTitleContainer.text
        soundTitle_usernameHeroContainer = driver.find_element(By.CLASS_NAME, "sc-link-secondary")
        varsoundTitle_usernameHeroContainer = soundTitle_usernameHeroContainer.text
        varLbl = varsoundTitle_usernameHeroContainer
        varsoundTitle_usernameHeroContainer = varsoundTitle_usernameHeroContainer.replace(' ', '')
        # playbackTimeline__duration
        ##playbackTimeline__duration =driver.find_element(By.CLASS_NAME,"playbackTimeline__duration")
        playbackTimeline__duration = driver.find_element(By.XPATH,
                                                         "*//div[contains(@class,'playbackTimeline__duration')]/span[2]")
        varplaybackTimeline_duration = playbackTimeline__duration.text
        varMin = varplaybackTimeline_duration.split(":")[0]
        varSec = varplaybackTimeline_duration.split(":")[1]
        lblGS = ttk.Label(canvas1).pack()
        ###lblGS = ttk.Label(canvas1).grid(column=1, row=11, sticky=('N'))#GRID
        canvas1.create_window(333,420,window=lblGS)#CREATE_WINDOW
        ttk.Label(canvas1, text="...gathered data from SoundCloud " + varStrTitle + " ...").pack()
        ###ttk.Label(canvas1, text="...gathered data from SoundCloud " + varStrTitle + " ...").grid(column=0, row=12, sticky='W')#GRID
        root.update()
        print('...gathered data from SoundCloud ' + varStrTitle + ' ...')

        varSheetname_GS = ''
        if varsoundTitle_usernameHeroContainer == 'FloatingBlueRecords' or varsoundTitle_usernameHeroContainer == 'DayDoseOfHouse':
            varSheetname_GS = varsoundTitle_usernameHeroContainer
        else:
            # look for sheetname as an input value entered by user
            varSheetname_GS = input("Enter the sheetname in (GooSheets):")
        vURL_GooShts = "https://docs.google.com/spreadsheets/d/1hpK3ziZq9QrdBi1FYnDy4SPKiciyo4UBwJBcHs0g2Rw/edit#gid=487846033"

        driver.quit()

        sheetname = varSheetname_GS
        sheet_id = "1hpK3ziZq9QrdBi1FYnDy4SPKiciyo4UBwJBcHs0g2Rw"
        xls = pd.ExcelFile(f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx')
        df = pd.read_excel(xls, sheetname, header=0)

        rowValues = df.loc[df['Release Title'] == varStrTitle]
        varIndxValueToCompare = df['Release Title'].loc[lambda x: x == varStrTitle].index
        varRT_to_write = rowValues['Release Title'].loc[rowValues.index[0]]

        numTire=varRT_to_write.count('-')
        if numTire >1:
            artistname_comma=varRT_to_write.split(",")
            if artistname_comma[-1].count('-')<=1:
                songTitle=artistname_comma[-1].split(" - ")[1]
                artNome0=",".join(artistname_comma[:-1])
                artNome=artNome0+','+artistname_comma[-1].split(" - ")[0]

                varArtistName = artNome
                varSongTitle = songTitle
            else:
                songTitle=artistname_comma[-1].split(" - ")[2]
                varSongTitle = songTitle
                artNome0='-'.join(artistname_comma[-1].split(" - ")[:-1])
                artNome=','.join(artistname_comma[:-1])+','+artNome0
                varArtistName=artNome
        else:
            varArtistName = varRT_to_write.split(" - ")[0]
            varSongTitle = varRT_to_write.split(" - ")[1]

        varG_to_write = rowValues['Genre'].loc[rowValues.index[0]]
        varUPC_to_write = rowValues['UPC'].loc[rowValues.index[0]]
        varPriority_to_write = rowValues['Priority'].loc[rowValues.index[0]]
        varRD_to_write = rowValues['d'].loc[rowValues.index[0]]
        if pd.isna(rowValues['d'].loc[rowValues.index[0]]):
            a = varIndxValueToCompare[0]
            for i in range(10):
                a = a - 1
                varRD_to_write = df['d'].loc[a]
                if pd.isna(varRD_to_write) == False:
                    break
        varRDF_to_write = '2022'
        if varSheetname_GS == 'DayDoseOfHouse':
            varRI_to_write = rowValues['Radio ISRC'].loc[rowValues.index[0]]
            varRE_to_write = rowValues['Radio/Extended'].loc[rowValues.index[0]]
            if varIndxValueToCompare < 449 and varIndxValueToCompare > 273:
                varRDF_to_write = varRD_to_write + ',2022'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2022'
            elif varIndxValueToCompare > 1 and varIndxValueToCompare < 273:
                varRDF_to_write = varRD_to_write + ',2021'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2021'
            elif  varIndxValueToCompare > 449:
                varRDF_to_write = varRD_to_write + ',2023'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2023'
        elif varSheetname_GS == 'FloatingBlueRecords':
            varRI_to_write = rowValues['ISRC'].loc[rowValues.index[0]]
            varRE_to_write = rowValues['COVER?'].loc[rowValues.index[0]]
            if varIndxValueToCompare > 4 and varIndxValueToCompare < 178:
                varRDF_to_write = varRD_to_write + ',2022'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2022'
            elif varIndxValueToCompare > 178 :
                varRDF_to_write = varRD_to_write + ',2023'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2023'
            elif varIndxValueToCompare <4:
                varRDF_to_write = varRD_to_write + ',2021'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2021'
        elif varSheetname_GS == 'DeepCloudMusic':
            varRI_to_write = rowValues['ISRC'].loc[rowValues.index[0]]
            varRE_to_write = rowValues['COVER?'].loc[rowValues.index[0]]
            if varIndxValueToCompare > 4 and varIndxValueToCompare <14:
                varRDF_to_write = varRD_to_write + ',2022'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2022'
            elif varIndxValueToCompare < 4:
                varRDF_to_write = varRD_to_write + ',2021'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2021'
            elif varIndxValueToCompare >14:
                varRDF_to_write = varRD_to_write + ',2022'
                varRDF_to_write = pd.to_datetime(varRDF_to_write).strftime('%d.%m.%Y')
                varRD_to_write = varRD_to_write + ', 2022'

        varCntry_to_write = rowValues['Country'].loc[rowValues.index[0]]
        allData_ColumnList = ['Artist Name', 'Song Title', 'Country', 'Release Date', 'Release Date (Formatted)', 'UPC',
                              'ISRC', 'Genre', 'Minute', 'Seconds', 'URL Link', 'Label Name', 'Priority', 'Radio/Extended']
        allData_RowList = [varArtistName, varSongTitle, varCntry_to_write, varRD_to_write, varRDF_to_write, varUPC_to_write,
                           varRI_to_write, varG_to_write, varMin, varSec, varUrl_SC, varLbl, varPriority_to_write,
                           varRE_to_write]
        allData_readyfordf = dict(zip(allData_ColumnList, allData_RowList))

        AllDataOutput = pd.DataFrame(allData_readyfordf, index=[0])

        lblES = ttk.Label(canvas1).pack()
        ###lblES = ttk.Label(canvas1).grid(column=1, row=13, sticky=('N'))#GRID
        ttk.Label(canvas1, text="...pulled out data also from googlesheets ...").pack()
        ###ttk.Label(canvas1, text="...pulled out data also from googlesheets ...").grid(column=0, row=14, sticky='W')#GRID
        root.update()
        print('...pulled out data also from googlesheets ...')

        def resource_path_output(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)

        def resolve_path(path):
            if getattr(sys, "frozen", False):
                # If the 'frozen' flag is set, we are in bundled-app mode!
                resolved_path = os.path.abspath(os.path.join(sys._MEIPASS, path))
                print(resolved_path,'-+-')
            else:
                # Normal development mode. Use os.getcwd() or __file__ as appropriate in your case...
                resolved_path = os.path.abspath(os.path.join(os.getcwd(), path))
                print(resolved_path, '+-+')

            return resolved_path


        if getattr(sys,'frozen',False) and hasattr(sys,'_MEIPASS'):
            print('runs in PyIns bundle')
            print(os.path.abspath)
            print(os.getcwd())
        else:
            print('runs in normal python process')
        new_path=os.getcwd()
        new_path02 = new_path[:-4]
        new_path03xl = new_path02 + 'outputFile/outputData.xlsm'

        with pd.ExcelWriter(new_path03xl, engine='openpyxl', mode='r+', if_sheet_exists='overlay',engine_kwargs={'keep_vba': True}) as writer:
            book = load_workbook(new_path03xl, keep_vba=True)

            print('read file')
            print(writer.engine)
            #writer.book = book
            #writer.sheets = {ws.title: ws for ws in book.worksheets}

            current_sheet = book['Sheet1']
            Column_A = current_sheet['A']
            maxrow = max(c.row for c in Column_A if c.value is not None)

            for sheetname in writer.sheets:
                AllDataOutput.to_excel(writer, sheet_name=sheetname, startrow=maxrow, index=False, header=False)

        ttk.Label(canvas1, text="...data retrieving finished. Check file ...").pack()
        root.update()

add_urlbox_button = tk.Button(canvas1, text="Add another URL", font='Segoe 9', command=addurlbox)
add_urlbox_button.pack(padx=3,pady=3)

button1 = tk.Button(text='scrape', command=crawler, bg='darkblue', fg='white',width=33)
button1.pack(padx=3,pady=3)

root.mainloop()
#pyinstaller -F --add-data "./outputFile/outputData.xlsm;./outputFile" SCtrackcrawlerv3.3.3.py --onefile --clean --add-binary "./driver/chromedriver.exe;./driver"
#pyinstaller SCtrackcrawlerv3.0.1.3.py --onefile --add-binary "./driver/chromedriver.exe;./driver" --add-data "./outputFile/outputData.xlsm;./outputFile"