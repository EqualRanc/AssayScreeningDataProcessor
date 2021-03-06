# -*- coding: utf-8 -*-
#"""
#This script is intended to automatically process raw UNMODIFIED .csv files read from the Envision
# to create a DB upload file.
#
# author: EqualRanc
#    
#The second part of this script is designed to handle and process data from SP screens.

import os
import PySimpleGUI as sg
import pandas as pd
import datetime
import xlwings as xw

#~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~

#Update tabno (27 in this present case) if new tabs are created. At this time of writing there were 27 tabs
#excluding the data 99 tab. The data 99 tab does not count for the first argument of processing oner (~line 343). 

tabno = 27
            
#~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~+*^*+~            
            


def process_xl(tabs, fullname):
    excel_app = xw.App(visible=False)
    filepath = os.path.expanduser(fullname)
    if not os.path.exists(filepath):
        return filepath
    excel_book = excel_app.books.open(filepath)
    sheetend = tabs + 1
    df = {}
    for number in range(1, sheetend):
        tabname = 'data ' + str(number)
        sheet = excel_book.sheets(tabname)
        df[tabname] = sheet[sheet.used_range.address].options(pd.DataFrame, index=False, header=True).value
        df[tabname] = df[tabname].fillna(0)
    lasttab = 'data 99'
    sheetlast = excel_book.sheets(lasttab)
    df[lasttab] = sheetlast[sheetlast.used_range.address].options(pd.DataFrame, index=False, header=True).value
    df[lasttab] = df[lasttab].fillna(0)
    excel_book.close()
    excel_app.quit()
    return df

def normalize384(foldername):
    path = os.path.expanduser(foldername)
    if not os.path.exists(path):
        return path
    ls = os.listdir(path)
    ls2 = []
    for temp in ls:
        if temp[-4:] == '.csv':
            if temp[:0] != '~':
                if temp[-14:] != 'DB_Upload.csv':
                    if temp[-18:] != 'Plate_Summary.xlsx':
                        ls2.append(temp)
    normalizedcoldata = []
    hisdata = []
    losdata = []
    hiadata = []
    loadata = []
    zdata = []
    windowdata = []
    for file in ls2:
        filename = path + '/' + file
        rawfile = pd.read_csv(filename, delimiter=',', header=None, skip_blank_lines=True, names=list(range(25)))
        rawfile.fillna(0)
        alocation = rawfile[0]=='A'
        plocation = rawfile[0]=='P'
        aloclist = rawfile.index[alocation].tolist()
        ploclist = rawfile.index[plocation].tolist()        
        samplesnp = rawfile.iloc[range(aloclist[0],ploclist[0]+1),range(1,23)].to_numpy(dtype=float)
        assayhi = rawfile.iloc[range(aloclist[0],aloclist[0]+8),24].to_numpy(dtype=float)
        assaylo = rawfile.iloc[range(aloclist[0]+8,ploclist[0]+1),24].to_numpy(dtype=float)
        hiavg =  assayhi.mean(axis=0)
        histd = assayhi.std(axis=0)
        loavg = assaylo.mean(axis=0)
        lostd = assaylo.std(axis=0)
        zcalc = 1-(3*((histd+lostd)/(hiavg-loavg)))
        windowcalc = hiavg/loavg
        normal = ((samplesnp - loavg) / (hiavg - loavg)) * 100
        normaldf = pd.DataFrame(normal)
        normalizedcoldata.append(list(normaldf.stack()))
        hisdata.append(histd)
        losdata.append(lostd)
        hiadata.append(hiavg)
        loadata.append(loavg)
        zdata.append(zcalc)
        windowdata.append(windowcalc)
    normalizedcoldata = [j for i in normalizedcoldata for j in i]
    return normalizedcoldata, hisdata, losdata, hiadata, loadata, zdata, windowdata

def normalize1536(foldername):
    path = os.path.expanduser(foldername)
    if not os.path.exists(path):
        return path
    ls = os.listdir(path)
    ls2 = []
    for temp in ls:
        if temp[-4:] == '.csv':
            if temp[:0] != '~':
                if temp[-13:] != 'DB_Upload.csv':
                    if temp[-18:] != 'Plate_Summary.xlsx':
                        ls2.append(temp)
    normalizedcoldata = []
    hisdata = []
    losdata = []
    hiadata = []
    loadata = []
    zdata = []
    windowdata = []
    for file in ls2:
        filename = path + '/' + file
        rawfile = pd.read_csv(filename, delimiter=',', header=None, skip_blank_lines=True, names=list(range(49)))
        alocation = rawfile[0]=='A'
        aflocation = rawfile[0]=='AF'
        aloclist = rawfile.index[alocation].tolist()
        afloclist = rawfile.index[aflocation].tolist()
        samplesnp = rawfile.iloc[range(aloclist[0],afloclist[0]+1),range(1,45)].to_numpy(dtype=float)
        assayhi = rawfile.iloc[range(aloclist[0],afloclist[0]+1),47].to_numpy(dtype=float)
        assaylo = rawfile.iloc[range(aloclist[0],afloclist[0]+1),48].to_numpy(dtype=float)
        hiavg =  assayhi.mean(axis=0)
        histd = assayhi.std(axis=0)
        loavg = assaylo.mean(axis=0)
        lostd = assaylo.std(axis=0)
        zcalc = 1-(3*((histd+lostd)/(hiavg-loavg)))
        windowcalc = hiavg/loavg
        normal = ((samplesnp - loavg) / (hiavg - loavg)) * 100
        normaldf = pd.DataFrame(normal)
        normaldftransposed = normaldf.T
        normalizedcoldata.append(list(normaldftransposed.stack()))
        hisdata.append(histd)
        losdata.append(lostd)
        hiadata.append(hiavg)
        loadata.append(loavg)
        zdata.append(zcalc)
        windowdata.append(windowcalc)
    normalizedcoldata = [j for i in normalizedcoldata for j in i]
    return normalizedcoldata, hisdata, losdata, hiadata, loadata, zdata, windowdata

def uinput():
    initialcheckboxlist = []
    if values['-SS-'] == True:
        initialcheckboxlist.append(["data 1", "data 2", "data 3"])
    if values['-S-'] == True:
        initialcheckboxlist.append(["data 4"])
    if values['-NN-'] == True:
        initialcheckboxlist.append(["data 5", "data 6", "data 7", "data 8"])
    if values['-N-'] == True:
        initialcheckboxlist.append(["data 9", "data 10", "data 11", "data 12"])
    if values['-O-'] == True:
        initialcheckboxlist.append(["data 13", "data 14"])
    if values['-AT-'] == True:
        initialcheckboxlist.append(["data 15"])
    if values['-CA-'] == True:
        initialcheckboxlist.append(["data 16", "data 17", "data 18", "data 19", "data 20", "data 21", "data 22", "data 23"])
    if values['-AH-'] == True:
        initialcheckboxlist.append(["data 24"])
    if values['-KT-'] == True:
        initialcheckboxlist.append(["data 25"])
    if values['-W-'] == True:
        initialcheckboxlist.append(["data 26", "data 27"])
    if values['-ORPH-'] == True:
        initialcheckboxlist.append(["data 99"])

    initialcheckboxlist = [j for i in initialcheckboxlist for j in i]
    if len(initialcheckboxlist) == 0:
        window['assaystatus'].print("Please pick at least one fragment class (Assay Tab).")
        
    #Filters out the data tabs you'd like to exclude from the assaycheckboxlist
    if len(values['-excl-']) != 0:
        exclinput = values['-excl-']
        
        #Converts user inputs into the tab names
        excl = ["data " + str(xin) for xin in exclinput]
        
        #Format for this line is [expression for item in iterable if condition == True]
        assaycheckboxlist = []
        assaycheckboxlist = [x1 for x1 in initialcheckboxlist if not any(x2 == x1 for x2 in excl)]
    else:
        assaycheckboxlist = initialcheckboxlist
    return assaycheckboxlist

def dfilter(oner, assaycheckboxlist):
    assayfullsheet = pd.concat(oner, keys=assaycheckboxlist)
    assayslice = assayfullsheet.loc[:, [
        'Class',
        'Molecule Name',
        'Batch Name',
        'Batch External Identifier',
        'Storage 96W or Box ID',
        'Well',
        'Concentration (mM)',
        '384W ML',
        '384W ML Well',
        '1536W ML',
        '1536W ML Well',
        '1536W LL',
        '1536W LL Well',
        '1536W ZI',
        '1536W ZI Well',
        '384W ZI',
        '384W ZI Well']]
    assayslice['Target'] = values['-target-']
    assayslice['Run Date'] = values['-rundate-']
    assayslice['Run ID'] = values['-runid-']
    assayslice['xb ID'] = values['-xbid-']
    assayslice['Assay ID'] = values['-assayid-']
    assayslice['Concentration (uM)'] = values['-conc-']
    return assayslice

def psummary(assayslice):
    zinames = []
    if values['-A1536-'] == True:
        assayslice = assayslice[(~assayslice['Molecule Name'].isin(['XB-01', 'BB8900', 'EC57', 'EC98']))]
        datalist, hisdata, losdata, hiadata, loadata, zdata, windowdata = normalize1536(foldername)
        assayslice['% Activity (DMSO) plate'] = datalist
        assayslice.drop('384W ZI',inplace=True,axis=1)
        assayslice.drop('384W ZI Well',inplace=True,axis=1)
        zinamestemp = []
        for tab in assaycheckboxlist:
            zinamestemp.append(oner[tab].loc[0, ['1536W ZI']])
        for series in zinamestemp:
            zinames.append(series[0])
        assayslice = assayslice[(assayslice['Class'] != 'empty')] #Cleans up DB upload sheet
    if values['-A384-'] == True:
        assayslice = assayslice[(~assayslice['Molecule Name'].isin(['xb', 'BB8900', 'EC57', 'EC98']))]
        zinamestemp=[]
        for tab in assaycheckboxlist:
            apnamestemp.append(oner[tab]['384W ZI'].unique())
        for array in zinamestemp:
            for item in array:
                zinames.append(item)
        zinames = pd.DataFrame(zinames,columns=['zinames'])
        if values['-CA-'] == True:
            assayslice = assayslice[(~assayslice['384W ZI'].isin(['ZI-92']))]
            zinames = zinames[(~zinames['zinames'].isin(['AP-92']))]
        if values['-S-'] == True:
            assayslice = assayslice[(~assayslice['384W ZI'].isin(['ZI-12']))]
            zinames = zinames[(~zinames['zinames'].isin(['AP-12']))]
        if values['-O-'] == True:
            assayslice = assayslice[(~assayslice['384W ZI'].isin(['ZI-56']))]
            zinames = zinames[(~zinames['zinames'].isin(['ZI-56']))]
        if values['-AT-'] == True:
            assayslice = assayslice[(~assayslice['384W ZI'].isin(['ZI-59', 'ZI-60']))]
            zinames = zinames[(~zinames['zinames'].isin(['ZI-59', 'ZI-60']))]
        zinames = list(zinames['zinames'])
        datalist, hisdata, losdata, hiadata, loadata, zdata, windowdata = normalize384(foldername)
        assayslice['% Activity (DMSO) plate'] = datalist
        assayslice.drop('1536W ZI',inplace=True,axis=1)
        assayslice.drop('1536W ZI Well',inplace=True,axis=1)
        assayslice = assayslice[(assayslice['Class'] != 'empty')] #Cleans up CDD upload sheet
    return assayslice, zinames, datalist, hisdata, losdata, hiadata, loadata, zdata, windowdata

def pexcel():
    psheaders = ["Plate Name","Fold Window","Z'","High Avg.","High Std. Dev.","Low Avg.","Low Std. Dev."]
    ps = xw.Book()
    pssheet = ps.sheets[0]
    pssheet.range('A1').value = psheaders
    pssheet.range('A2').options(transpose=True).value = zinames
    pssheet.range('B2').options(transpose=True).value = windowdata
    pssheet.range('C2').options(transpose=True).value = zdata
    pssheet.range('D2').options(transpose=True).value = hiadata
    pssheet.range('E2').options(transpose=True).value = hisdata
    pssheet.range('F2').options(transpose=True).value = loadata
    pssheet.range('G2').options(transpose=True).value = losdata
    return ps

def datalayout():            
    #Define checkboxes for fragment classes
    monocheckboxes = [[sg.Checkbox(':XA', default=False, key="-NN-"), sg.Checkbox(':XB', default=False, key="-N-"),
                       sg.Checkbox(':1Z', default=False, key="-S-"), sg.Checkbox(':Z', default=False, key="-SS-"),
                       sg.Checkbox(':O', default=False, key="-O-"), sg.Checkbox(':Q', default=False, key="-AT-"),
                       sg.Checkbox(':ZO', default=False, key="-CA-"), sg.Checkbox(':TGTT', default=False, key="-AH-"),
                       sg.Checkbox(':BR', default=False, key="-KT-"), sg.Checkbox(':W', default=False, key="-W-"), sg.Checkbox(':ID', default=False, key="-ORPH-")]
    ]

    #Setup for the singple-point data processing main interface
    dataentries = [
        [sg.Text("Single-Point Assay Data Processing Tool",font='Any 18')],
        [sg.Frame('Browse to raw screening data folder:', [[sg.Input(key='-rawdata-'), sg.FolderBrowse(target='-rawdata-')]])],
        [sg.Frame('Browse to the chemical database file:', [[sg.Input(key='-oner-'), sg.FileBrowse(target='-oner-')]])],
        [sg.Frame("Please select desired assay plate type:", [[sg.Radio('1536W Assay', "AssayType", default=False, key="-A1536-"), sg.Radio('384W Assay', "AssayType", default=False, key="-A384-")]])],
        [sg.Frame("Choose among the following classes:", monocheckboxes)],
        [sg.Frame("Enter metadata, e.g. xb ID, LL type, assay info, target, assay concentration:",
               [
                   [
                       sg.Text("Run Date:"),sg.Input(key="-rundate-"), sg.Text("Run ID:"),sg.Input(key="-runid-")],
                   [
                       sg.Text("xb ID:"),sg.Input(key="-xbid-"), sg.Text("Target"),sg.Input(key="-target-")],
                   [
                       sg.Text("Assay ID:"),sg.Input(key="-assayid-"), sg.Text("Concentration (uM):"),sg.Input(key="-conc-")]
               ]
              )
         ],
        [sg.Frame("Enter any assay plate numbers you'd like to exclude:", [[sg.Input(key='-excl-')]])],
        [sg.Submit(), sg.Cancel()]
    ]


    datastatus = [
        [sg.Text('Status:', size=[20,1])],
        [sg.Multiline(key='datastatus',autoscroll=True,size=(30,20))],
    ]

    datalayout = [
        [
        sg.Column(dataentries),
        sg.VSeperator(),
        sg.Column(datastatus)
        ]
    ]
    return datalayout

# Creates the theme
sg.theme('Dark Teal')
datalayout = datalayout()

# Creates the window
window = sg.Window('Single-Point Assay Data Processing Tool', datalayout, no_titlebar=False, alpha_channel=.9, grab_anywhere=True)


# Create event loop to enable user inputs
while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Cancel"):
        break
    
    # Processing steps for renaming assay plates tab
    elif event == "Submit":
        try:
            foldername = values['-rawdata-']
            cdb = values['-oner-']
            rundate = values["-rundate-"]
            runid = values["-runid-"]
            xbid = values["-xbid-"]
            assayid = values["-assayid-"]
            conc = values["-conc-"]
            
            oner = process_xl(tabno, cdb)

            # Processes user input
            try:
                assaycheckboxlist = uinput()
            except:
                window['datastatus'].print("Number of raw files does not match chemical class and assay plate exclusions selected.")
                break
                
            #Filters One Ring tabs and joins them into one table based on fragment class selections, creates metadata columns
            assayslice = dfilter(oner, assaycheckboxlist)
            
            #Prepares plate summary sheet
            assayslice, apnames, datalist, hisdata, losdata, hiadata, loadata, zdata, windowdata = psummary(assayslice)
                
            #Creates the plate summary xlsx file
            ps = pexcel()
            
            #Prepare to export the .csv and .xlsx files            
            if values['-A1536-'] == True:
                assayfiletype = '1536W'
            elif values['-A384-'] == True:
                assayfiletype = '384W'
            else:
                assayfiletype = 'Unknown Plate Type'
            outfile = str(values['-rawdata-']) + '/' + str(datetime.date.today().isoformat()) + '_' + '%s' % str(values['-xbid-']) + '_' + str(assayfiletype) + '_' + values['-assayid-'] + '_' + 'DB_Upload.csv'
            assayslice.to_csv(outfile, index=False)
            outfile2 = str(values['-rawdata-']) + '/' + str(datetime.date.today().isoformat()) + '_' + '%s' % str(values['-xbid-']) + '_' + str(assayfiletype) + '_' + values['-assayid-'] + '_' + 'Plate_Summary.xlsx'
            ps.save(path=outfile2)


            window['datastatus'].print("Processing of single-point raw files complete.")
        except:
            window['datastatus'].print("Unable to process assay raw files.")
window.close()
