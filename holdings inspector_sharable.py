# Preliminary setup stuff:
#(Loading libraries, setting up dictionaries, asking for output filename, etc.)

import requests
import pandas as pd
from bs4 import BeautifulSoup
import re

#Define empty dictionary for later
dict = {}
dict['mmsid'] = []
dict['holdingid'] = []
dict['location'] = []
dict['852firstindicator'] = []
dict['852secondindicator'] = []
dict['callnumber'] = []
dict['fullsubfieldh'] = []
dict['subfieldicount'] = []
dict['subfieldhspace'] = []
dict['badsubfieldhpattern'] = []
dict['subfieldk'] = []
dict['subfieldl'] = []

#Ask for an output filename for later use
outputname = input('Enter output filename without extension: ')


#Set up default things for base URL and bib API key
alma_base = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1'
#IZ Bibs read-only API key
bibapi = 'INSERT YOUR BIBS API KEY HERE'
headers = {"Accept": "application/xml"}

# Next section takes the input file and prepares it for use
#Note: I keep the MMS ID and Holdings ID in ""s until ready to use them so there's no chance of them getting messed up. The script removes those extra ""s.

inputname = input('MMSID Input Filename without extension: ')
source = pd.read_excel(f"{inputname}.xlsx")
lookupmms_raw=source['MMS Id'].astype(str)
lookupholdid_raw=source['Holding Id'].astype(str)
lookupcallno=source['Permanent Call Number'].astype(str)

#Removes preventative ""s from MMS and Holdings ID information
lookupmms = ([s.replace('"', '') for s in lookupmms_raw])
lookupholdid = ([t.replace('"', '') for t in lookupholdid_raw])

#Print just to make sure they look ok
print(lookupmms[0])
print(lookupholdid[0])
print(lookupcallno[0])

# Runs the loop to look up records

for i, mms in enumerate(lookupmms, 0):
    #sleep(randint(5,100))    

    #Pulls in MMS and Holdings ID from input data, and appends those and the call number provided by analytics and puts those in the dictionary to keep things in order:
   
    mms = lookupmms[i]
    dict['mmsid'].append(mms)
    
    holdid = lookupholdid[i]
    dict['holdingid'].append(holdid)
    
    callno = lookupcallno[i]
    dict['callnumber'].append(callno)
    
    #Print statement to let you know how many it's done and how much is left to do
    print(f"Processing {mms}, entry # {str(i+1)} of {str(len(lookupmms))}")

    #The API request itself
    r = requests.get(f"{alma_base}/bibs/{mms}/holdings/{holdid}?apikey={bibapi}", headers=headers)

    # Creating the Soup Object containing all data
    soup = BeautifulSoup(r.content, "xml")
    
    #Grab indicators from 852 field
    f852 = soup.find("datafield", attrs={"tag": '852'})
    
    try:
        firstind = f852['ind1']
        print(firstind)
        dict['852firstindicator'].append(firstind)
    except TypeError:
        firstind = "error"
        print(firstind)
        dict['852firstindicator'].append(firstind)
    
    try:
        secondind = f852['ind2']
        print(secondind)
        dict['852secondindicator'].append(secondind)
    except TypeError:
        secondind = "error"
        print(secondind)
        dict['852secondindicator'].append(secondind)

    #Get location code from 852 $c:
    try: 
        #Grab location from subfield c in 852
        subfieldc = f852.find("subfield", attrs={"code": "c"}).getText()
        print(subfieldc)    
        dict['location'].append(subfieldc)
    except AttributeError:
        subfieldc = "error"
        print(subfieldc)    
        dict['location'].append(subfieldc)
        
    #Get the 852 $h    
    try: 
        #Grab just subfield h in 852
        subfieldh = f852.find("subfield", attrs={"code": "h"}).getText()
        print(subfieldh)    
        dict['fullsubfieldh'].append(subfieldh)
        
        #Does $h start with spaces?
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

        #Are there extra spaces in the $h?
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
        
    except AttributeError:
        subfieldh = ""
        print(subfieldh)    
        dict['fullsubfieldh'].append(subfieldh)
        subfieldhspaces = ""
        dict['subfieldhspace'].append(subfieldhspaces)
        subfieldhpattern = ""
        dict['badsubfieldhpattern'].append(subfieldhpattern)


    #Find out how many subfield Is there are in the 852
    try:
        findsubfieldi = f852.findAll("subfield", attrs={"code": "i"})
        howmanyIs = len(findsubfieldi)
        print(howmanyIs)
        dict['subfieldicount'].append(howmanyIs)
    except AttributeError:
        howmanyIs = 0
        dict['subfieldicount'].append(howmanyIs)
        
 
    #Look to see if there are any 852 $k (call number prefix)    
    try: 
        #Get $k if present
        subfieldk = f852.find("subfield", attrs={"code": "k"}).getText()
        print(subfieldk)    
        dict['subfieldk'].append(subfieldk)
    except AttributeError:
        subfieldk = ""
        print(subfieldk)    
        dict['subfieldk'].append(subfieldk)

    #Look to see if there is already a $l in the 852 (and if so, what)    
    try: 
        #Grab just subfield l in 852
        subfieldl = f852.find("subfield", attrs={"code": "l"}).getText()
        print(subfieldl)    
        dict['subfieldl'].append(subfieldl)
    except AttributeError:
        subfieldl = ""
        print(subfieldl)    
        dict['subfieldl'].append(subfieldl)


# Takes retrieved data and turns it into a useable Excel spreadsheet:

#Turn dictionary with all of the data into a dataframe
df = pd.DataFrame(dict)
df.head()

#Exports dataframe to Excel using filename entered at top of script
df.to_excel(f'{outputname}.xlsx', index=None)