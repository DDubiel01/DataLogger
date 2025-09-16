import pandas as pd
import pygsheets

def import_data(credentials, book_name, data_sheet, whitelist_sheet):
    ''' 
    Inputs:

    credentials: path to google credentials json
    book_name: str; name of workbook
    data_sheet: str; name of sheet to pull data from
    whitelist_sheet: str; name of sheet with the white list decoder

    Outputs (see docs for frames):

    rawdf, names_dict (dictionary of code name and plain name)
    '''

    google_sheet = pygsheets.authorize(credentials = credentials)
    google_book = google_sheet.open(book_name)
    rawdata = google_book.worksheet('title',data_sheet).get_all_records(head = 1)
    decode = google_book.worksheet('title',whitelist_sheet).get_all_records(head = 1)


    #Make Dataframes
    rawdf = pd.DataFrame(rawdata)
    decodedf = pd.DataFrame(decode)

    #Make some things to hold on to
    names_raw = decodedf['Code']
    names_plain = decodedf['Name']
    names_dict = dict(zip(decodedf['Code'],decodedf['Name'])) 

    return [rawdf, names_dict]

def create_basicframe(rawdf, names_dict):
    ''' 
    Inputs:

    rawdf
    names_dict

    Outputs (see docs for frames):

    basicframe
    '''

    #Get Snipe/Sniped Totals
    sniper_totals = rawdf['Sniper'].value_counts()
    sniped_totals = rawdf['Sniped'].value_counts()

    #Initialize
    basicframe = pd.DataFrame(0,index = list(names_dict.keys()), columns = ['Snipes','Sniped', 'K/D'])

    #Populate the Frame
    for idx, val in sniper_totals.items():
        basicframe.at[idx, 'Snipes'] = val

    for idx, val in sniped_totals.items():
        basicframe.at[idx, 'Sniped'] = val

    #Calculate KDs
    basicframe['K/D'] = basicframe['Snipes'] / basicframe['Sniped']

    basicframe.rename(index = names_dict, inplace = True)

    return basicframe

def create_comboframe(rawdf, names_dict):
    ''' 
    Inputs:

    rawdf
    names_dict

    Outputs (see docs for frames):

    comboframe
    '''
    #Initialize the combo frame
    comboframe=pd.DataFrame(0,index=list(names_dict.keys()),columns=list(names_dict.keys()))

    #iterate
    for rownum, row in rawdf.iterrows():
        #comboframe.at[name of sniper, name of sniped] = that location + 1
        comboframe.at[row.iloc[0],row.iloc[1]] = comboframe.at[row.iloc[0],row.iloc[1]] + 1
    
    return comboframe

def create_uniqueframe(comboframe: pd.DataFrame, names_dict):
    ''' 
    Inputs:

    comboframe
    names_dict

    Outputs (see docs for frames):

    uniqueframe
    '''

    #Initialize the Unique Frame
    uniqueframe = pd.DataFrame(0, index = list(names_dict.keys()), columns = ['No. Diff People Sniped'])

    for row, rownum in uniqueframe.iterrows():
        #For each name, go through their row in the comboframe, and count the number of times the number is greater than 0
        uniqueframe.loc[row, 'No. Diff People Sniped'] = sum(comboframe.loc[row]>0)
        
    uniqueframe.rename(index = names_dict, inplace = True)

    return uniqueframe

def create_timetotal(rawdf):
    ''' 
    Inputs:

    rawdf

    Outputs (see docs for frames):

    dateframe
    weekframe
    truncdateframe
    '''
    #Truncate the dates into month/day/year
    truncdateframe = rawdf
    truncdateframe['Date'] = pd.to_datetime(truncdateframe['Date'])

    for rownum, row in truncdateframe.iterrows():
        truncdateframe.at[rownum,'Date'] = truncdateframe.at[rownum,'Date'].strftime('%m/%d/%Y')
        truncdateframe.at[rownum,'Day of Week'] = truncdateframe.at[rownum,'Date'].weekday()
        
    #Make a date range between the first and last day
    daterange = pd.date_range(truncdateframe['Date'][0], truncdateframe['Date'].iloc[-1])
    #Make an empty data frame with a row for each day
    dateframe = pd.DataFrame(index = daterange, columns = ['Snipes'])
    #Count the frequency of each date
    unsorted = truncdateframe['Date'].value_counts()

    #Iterate to input the date frequency into the dateframe
    for row, rownum in dateframe.iterrows():
        #value_counts() wont generate a value for dates that have no frequency
        try:
            unsorted[row]
        except:
            dateframe.loc[row]['Snipes'] = 0
        else:
            dateframe.loc[row]['Snipes'] = unsorted[row]
            
    #rearrange the dateframe based on weeks
    weekframe = dateframe.resample('W-SUN', closed = 'left', label = 'left').sum()

    return dateframe, weekframe, truncdateframe, daterange

def create_timeind(truncdateframe,names_dict, daterange):
    ''' 
    Inputs:

    truncdateframe
    names_dict
    daterange: pandas datetime index of dates that had snipes

    Outputs (see docs for frames):

    sn_dateframe
    sd_dateframe
    sn_byweek
    sd_byweek
    '''

    #Initialize. These frames track the number of snipes per day by person
    sn_dateframe = pd.DataFrame(0, index = daterange, columns = list(names_dict.keys()))
    sd_dateframe = pd.DataFrame(0, index = daterange, columns = list(names_dict.keys()))

    #Iterate through the truncdateframe
    for rownum, row in truncdateframe.iterrows():
        #add one to the location (Day of the snipe, person who sniped/was sniped)
        sn_dateframe.at[truncdateframe.at[rownum,'Date'],truncdateframe.at[rownum,'Sniper']] += 1
        sd_dateframe.at[truncdateframe.at[rownum,'Date'],truncdateframe.at[rownum,'Sniped']] += 1
        
    sn_dateframe.rename(columns = names_dict, inplace = True)
    sd_dateframe.rename(columns = names_dict, inplace = True)

    #resort into weeks
    sn_byweek = sn_dateframe.resample('W-SUN', closed = 'left', label = 'left').sum()
    sd_byweek = sd_dateframe.resample('W-SUN', closed = 'left', label = 'left').sum()

    sn_byweek.rename(columns = names_dict, inplace = True)
    sd_byweek.rename(columns = names_dict, inplace = True)

    return sn_dateframe, sd_dateframe, sn_byweek, sd_byweek

def create_firstframe(names_dict, sn_dateframe, sd_dateframe):
    ''' 
    Inputs:
    
    names_dict
    sn_dateframe
    sd_dateframe

    Outputs (see docs for frames):

    firstframe
    '''

    firstframe = pd.DataFrame(index = list(names_dict.values()), columns = ['First Snipe Date', 'Number of Sniping Days', 'First Date Sniped', 'Number of Sniped Days'])

    for name in list(names_dict.values()):
        #First, check to see if the person has recorded a snipe
        try:
            #the nonzero() method puts out a list of integers of the days with non zero snipes for each name
            firstframe.loc[name,'First Snipe Date'] = sn_dateframe.index.values[sn_dateframe[name].to_numpy().nonzero()[0][0]]
        
        except IndexError:
            #If there are no days with non zero snipes, pass
            pass
        
        else:
            #The length of the nonzero() list wil be the days someone has recorded a snipe
            firstframe.loc[name,'Number of Sniping Days'] = len(sn_dateframe[name].to_numpy().nonzero()[0])
                    
        try:
            firstframe.loc[name,'First Date Sniped'] = sd_dateframe.index.values[sd_dateframe[name].to_numpy().nonzero()[0][0]]
        
        except IndexError:
            pass
        
        else:
            firstframe.loc[name,'Number of Sniped Days'] = len(sd_dateframe[name].to_numpy().nonzero()[0])
            
    return firstframe
            

