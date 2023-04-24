import requests
import pandas as pd
from bs4 import BeautifulSoup
import re

#Set up functions for each possible thing to look for (to be toggled on and off later): 
def locationcode (f852):
    try: 
        #Grab location from subfield c in 852
        subfieldc = f852.find("subfield", attrs={"code": "c"}).getText()
        print(subfieldc)    
        dict['location'].append(subfieldc)
    except AttributeError:
        subfieldc = "error"
        print(subfieldc)    
        dict['location'].append(subfieldc)
    return subfieldc

def firstindicator (f852):
    try:
        firstind = f852['ind1']
        print(firstind)
        dict['852firstindicator'].append(firstind)
    except TypeError:
        firstind = "error"
        print(firstind)
        dict['852firstindicator'].append(firstind)
    return firstind

def secondindicator (f852):
    try:
        secondind = f852['ind2']
        print(secondind)
        dict['852secondindicator'].append(secondind)
    except TypeError:
        secondind = "error"
        print(secondind)
        dict['852secondindicator'].append(secondind)
    return secondind

def findsubfieldh (f852):
    try: 
        #Grab just subfield h in 852
        subfieldh = f852.find("subfield", attrs={"code": "h"}).getText()
        print(subfieldh)    
        dict['fullsubfieldh'].append(subfieldh)
    except AttributeError:
        subfieldh = ""
        print(subfieldh)    
        dict['fullsubfieldh'].append(subfieldh)
    return(subfieldh)
        
def initialsubfieldhspace (f852):
    try:
            findsubfieldh = f852.find("subfield", attrs={"code": "h"}).getText()
            subfieldhspacesmatch = re.match(r'^\s', findsubfieldh)
            if subfieldhspacesmatch is not None:
                subfieldhspaces = True
            else:
                subfieldhspaces = False      
            print(subfieldhspaces)
            dict['subfieldhspace'].append(subfieldhspaces)
    except AttributeError:
        subfieldhspaces = "Nothing found"
        dict['subfieldhspace'].append(subfieldhspaces)
    return(subfieldhspaces)

def othersubfieldhspace (f852):
    #Are there extra spaces in the $h?
        findsubfieldh = f852.find("subfield", attrs={"code": "h"}).getText()
        try:
            subfieldhpatternmatch = re.match(r'([A-Z]+\s)', findsubfieldh)
            if subfieldhpatternmatch is not None:
                subfieldhpattern = True
            else:
                subfieldhpattern = False
            print(subfieldhpattern)
            dict['badsubfieldhpattern'].append(subfieldhpattern)
        except AttributeError:
            subfieldhpattern = "Nothing found"
            dict['badsubfieldhpattern'].append(subfieldhpattern)

def countsubfieldi (f852):
    #Find out how many subfield Is there are in the 852
    try:
        findsubfieldi = f852.findAll("subfield", attrs={"code": "i"})
        howmanyIs = len(findsubfieldi)
        print(howmanyIs)
        dict['subfieldicount'].append(howmanyIs)
    except AttributeError:
        howmanyIs = 0
        dict['subfieldicount'].append(howmanyIs)
    return(howmanyIs)

def findsubfieldk (f852):
    try: 
        #Get $k if present
        subfieldk = f852.find("subfield", attrs={"code": "k"}).getText()
        print(subfieldk)    
        dict['subfieldk'].append(subfieldk)
    except AttributeError:
        subfieldk = ""
        print(subfieldk)    
        dict['subfieldk'].append(subfieldk)
    return(subfieldk)

def findsubfieldl (f852):
    try: 
        #Grab just subfield l in 852
        subfieldl = f852.find("subfield", attrs={"code": "l"}).getText()
        print(subfieldl)    
        dict['subfieldl'].append(subfieldl)
    except AttributeError:
        subfieldl = ""
        print(subfieldl)    
        dict['subfieldl'].append(subfieldl)
    return(subfieldl)

#Set up default things for base URL and bib API key
alma_base = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1'
#IZ Bibs read-only API key
bibapi = 'INSERT YOUR BIBS API KEY HERE'
headers = {"Accept": "application/xml"}


inputname = input('MMSID Input Filename without extension: ')
source = pd.read_excel(inputname + '.xlsx')
lookupmms_raw=source['MMS Id'].astype(str)
lookupholdid_raw=source['Holding Id'].astype(str)
lookupcallno=source['Permanent Call Number'].astype(str)

lookupmms = ([s.replace('"', '') for s in lookupmms_raw])
lookupholdid = ([t.replace('"', '') for t in lookupholdid_raw])

print(lookupmms[0])
print(lookupholdid[0])
print(lookupcallno[0])

#Ask for an output filename for later use
outputname = input('Enter output filename without extension: ')

#Set up dictionary with constant data
dict = {}
dict['mmsid'] = []
dict['holdingid'] = []
dict['callnumber'] = []

#Ask which things to run and set up dictionary slots for ones that will run:
yes_choices = ['yes', 'y']
##no_choices = ['no', 'n']

selectfirstindicator = input('Report on first indicator? (yes/no): ')
if selectfirstindicator.lower() in yes_choices:
    runfirstindicator = True
    dict['852firstindicator'] = []
else:
    runfirstindicator = False

selectsecondindicator = input('Report on second indicator? (yes/no): ')
if selectsecondindicator.lower() in yes_choices:
    runsecondindicator = True
    dict['852secondindicator'] = []
else:
    runsecondindicator = False

selectlocation = input('Report on location code (852 $c)? (yes/no): ')
if selectlocation.lower() in yes_choices:
    runlocation = True
    dict['location'] = []
else:
    runlocation = False

selectfullsubfieldh = input('Report on full 852 $h? (yes/no): ')
if selectfullsubfieldh.lower() in yes_choices:
    runfullsubfieldh = True
    dict['fullsubfieldh'] = []
else:
    runfullsubfieldh = False

selectfieldhspace = input('Report on any spaces at start of 852 $h? (yes/no): ')
if selectfieldhspace.lower() in yes_choices:
    runfieldhspace = True
    dict['subfieldhspace'] = []
else:
    runfieldhspace = False

selectbadsubfieldhpattern = input('Report on any spaces in the middle of 852 $h? (yes/no): ')
if selectbadsubfieldhpattern.lower() in yes_choices:
    runbadsubfieldhpattern = True
    dict['badsubfieldhpattern'] = []
else:
    runbadsubfieldhpattern = False

selectsubfieldicount = input('Count occurences of 852 $i? (yes/no): ')
if selectsubfieldicount.lower() in yes_choices:
    runsubfieldicount = True
    dict['subfieldicount'] = []
else:
    runsubfieldicount = False

selectsubfieldk = input('Report on 852 $k? (yes/no): ')
if selectsubfieldk.lower() in yes_choices:
    runsubfieldk = True
    dict['subfieldk'] = []
else:
    runsubfieldk = False

selectsubfieldl = input('Report on 852 $l? (yes/no): ')
if selectsubfieldl.lower() in yes_choices:
    runsubfieldl = True
    dict['subfieldl'] = []
else:
    runsubfieldl = False


#Main lookup loop, will run functions based on selections made above
for i, mms in enumerate(lookupmms, 0): 
       
    mms = lookupmms[i]
    dict['mmsid'].append(mms)
    print(f"Processing {mms}, entry # {str(i+1)} of {str(len(lookupmms))}")
    
    holdid = lookupholdid[i]
    dict['holdingid'].append(holdid)
    
    callno = lookupcallno[i]
    dict['callnumber'].append(callno)
    
    r = requests.get(f"{alma_base}/bibs/{mms}/holdings/{holdid}?apikey={bibapi}", headers=headers)

    # Creating the Soup Object containing all data
    soup = BeautifulSoup(r.content, "xml")
    
    #Grab indicators from 852 field
    f852 = soup.find("datafield", attrs={"tag": '852'})
    
    #If function was selected by prompt, run, otherwise, skip that function and keep going:
    if runfirstindicator == True: 
        firstind = firstindicator(f852)
    else:
        pass

    if runsecondindicator == True:
        secondind = secondindicator(f852)
    else:
        pass

    if runlocation == True:
        subfieldc = locationcode(f852)
    else:
        pass

    if runfullsubfieldh == True:
        subfieldh = findsubfieldh(f852)
    else:
        pass
    
    if runfieldhspace == True:
        subfieldhspaces = initialsubfieldhspace(f852)
    else:
        pass
    
    if runbadsubfieldhpattern == True:
        subfieldhpattern = othersubfieldhspace(f852)
    else:
        pass

    if runsubfieldicount == True:
        howmanyIs = countsubfieldi(f852)
    else:
        pass
    
    if runsubfieldk == True:
        subfieldk = findsubfieldk(f852)    
    else:
        pass
    
    if runsubfieldl == True:
        subfieldl = findsubfieldl (f852)
    else:
        pass

#Turn dictionary with all of the data into a dataframe
df = pd.DataFrame(dict)
df.head()

#Exports dataframe to Excel using filename entered at top of script
df.to_excel(f'{outputname}.xlsx', index=None)
