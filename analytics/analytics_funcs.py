import pandas as pd
import pygsheets
import matplotlib.pyplot as plt
import numpy as np

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
    daterange
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
            
def plot_weekdaypie(truncdateframe, figlength = 10, figheight = 5, titlesize = 18):
    ''' 
    Inputs:

    truncdateframe
    kwarg figure size parameters

    Outputs (see docs for frames):

    weekdaypie
    '''

    sumsbyweekday = truncdateframe['Day of Week'].value_counts()
    sumsbyweekday.rename(index = { 0.0 : 'Sunday',
                                1.0 : 'Monday',
                                2.0 : 'Tuesday',
                                3.0 : 'Wednesday',
                                4.0 : 'Thursday',
                                5.0 : 'Friday',
                                6.0 : 'Saturday'}, inplace = True)

    weekdaypie = sumsbyweekday.plot.pie(figsize = (figlength, figheight),
                        ylabel = '')

    weekdaypie.set_title('Snipes per Day of the Week', fontsize = titlesize)
   
    return weekdaypie

def plot_overtime(dateframe, figlength = 10, figheight = 5, titlesize = 18, xsize = 14, ysize = 14):
    ''' 
    Inputs:

    dateframe
    kwarg figure size parameters

    Outputs (see docs for frames):

    overtime plot
    '''

    overtime_fig, overtime_ax = plt.subplots()

    y = dateframe['Snipes'].cumsum().values
        
    overtime_ax.plot(dateframe.index.values, y)

    overtime_ax.grid(axis = 'y')

    overtime_fig.set_figwidth(figlength)
    overtime_fig.set_figheight(figheight)

    overtime_ax.set_xlabel('Date', fontsize = xsize)
    overtime_ax.set_ylabel('Snipes', fontsize = ysize)
    overtime_ax.set_title('Snipes Over Time', fontsize = titlesize)

    return overtime_fig

def plot_topSn_overtime(basicframe, sn_dateframe, topqty, figlength = 10, figheight = 5, titlesize = 18, xsize = 14, ysize = 14, legendsize = 10):
    '''
    Inputs:

    basicframe
    sn_dateframe
    topqty (int): number of top snipers to plot
    kwarg figure size parameters

    Outputs (see docs for frames):

    plot of top snipers over time
    '''

    #Get Names of the top snipers
    topsns = basicframe.sort_values('Snipes', ascending = False).head(topqty).index.values
    #Make a dataframe without the top snipers
    noleaders_sn_dateframe = sn_dateframe.loc[:, ~sn_dateframe.columns.isin(topsns)]
    #Turn that frame into a single column
    noleaders_sn_dateframe['Sum'] = noleaders_sn_dateframe.sum(axis = 1)
    #Get a frame of just the top snipers
    onlyleaders_sn_dateframe = sn_dateframe.loc[:,topsns]

    #Create a column calculating the average snipes each day
    noleaders_sn_dateframe['Average'] = noleaders_sn_dateframe.loc[:, noleaders_sn_dateframe.columns != 'Sum'].mean(axis = 1)
    #Plot

    snleaderplot_fig, snleaderplot_ax = plt.subplots()
    snleaderframe = onlyleaders_sn_dateframe.join(noleaders_sn_dateframe['Average'])

    snleaderframe.cumsum().plot(figsize = (figlength, figheight), ax = snleaderplot_ax)

    snleaderplot_ax.grid(axis = 'y')

    snleaderplot_fig.set_figwidth(figlength)
    snleaderplot_fig.set_figheight(figheight)

    snleaderplot_ax.legend(fontsize = legendsize)

    snleaderplot_ax.set_xlabel('Date', fontsize = xsize)
    snleaderplot_ax.set_ylabel('Snipes', fontsize = ysize)
    snleaderplot_ax.set_title('Top {} Snipers'.format(topqty), fontsize = titlesize)

    return snleaderplot_fig

def plot_topSn_pie(basicframe, sn_dateframe, topqty, figlength = 10, figheight = 5, titlesize = 18):
    '''
    Inputs:

    dateframe
    kwarg figure size parameters

    Outputs (see docs for frames):

    overtime plot
    '''
    #Get Names of the top snipers
    topsns = basicframe.sort_values('Snipes', ascending = False).head(topqty).index.values 
    #Make a dataframe without the top snipers
    noleaders_sn_dateframe = sn_dateframe.loc[:, ~sn_dateframe.columns.isin(topsns)]
    #Turn that frame into a single column
    noleaders_sn_dateframe['Sum'] = noleaders_sn_dateframe.sum(axis = 1)
    #Get a frame of just the top snipers
    onlyleaders_sn_dateframe = sn_dateframe.loc[:,topsns]

    #Create a column calculating the average snipes each day
    noleaders_sn_dateframe['Average'] = noleaders_sn_dateframe.loc[:, noleaders_sn_dateframe.columns != 'Sum'].mean(axis = 1)
    #Plo
    sntopandothers = onlyleaders_sn_dateframe.join(noleaders_sn_dateframe['Sum'])
    sntopandothers.rename(columns = {'Sum':'Everyone Else'}, inplace = True)
    top5pie_sn_ax = sntopandothers.sum(axis = 0).plot.pie(figsize = (figlength, figheight), autopct='%1.1f%%')

    top5pie_sn_ax.set_title('Top {} Snipers'.format(topqty), fontsize = titlesize)
    return top5pie_sn_ax

def plot_topSd_overtime(basicframe, sd_dateframe, topqty, figlength = 10, figheight = 5, titlesize = 18, xsize = 14, ysize = 14, legendsize = 10):
    '''
    Inputs:

    basicframe
    sd_dateframe
    topqty (int): number of top snipers to plot
    kwarg figure size parameters

    Outputs (see docs for frames):

    Chart of top sniped over time
    '''

    #Get Names of the top snipers
    topsds = basicframe.sort_values('Sniped', ascending = False).head(topqty).index.values
    #Make a dataframe without the top snipers
    noleaders_sd_dateframe = sd_dateframe.loc[:, ~sd_dateframe.columns.isin(topsds)]
    #Turn that frame into a single column
    noleaders_sd_dateframe['Sum'] = noleaders_sd_dateframe.sum(axis = 1)
    #Get a frame of just the top snipers
    onlyleaders_sd_dateframe = sd_dateframe.loc[:,topsds]

    #Create a column calculating the average snipes each day
    noleaders_sd_dateframe['Average'] = noleaders_sd_dateframe.loc[:, noleaders_sd_dateframe.columns != 'Sum'].mean(axis = 1)
    #Plot

    sdleaderplot_fig, sdleaderplot_ax = plt.subplots()
    sdleaderframe = onlyleaders_sd_dateframe.join(noleaders_sd_dateframe['Average'])

    sdleaderframe.cumsum().plot(figsize = (figlength, figheight), ax = sdleaderplot_ax)

    sdleaderplot_ax.grid(axis = 'y')

    sdleaderplot_ax.legend(fontsize = legendsize)

    sdleaderplot_fig.set_figwidth(figlength)
    sdleaderplot_fig.set_figheight(figheight)

    sdleaderplot_ax.set_xlabel('Date', fontsize = xsize)
    sdleaderplot_ax.set_ylabel('Sniped', fontsize = ysize)
    sdleaderplot_ax.set_title('Top 5 Sniped', fontsize = titlesize)

    return sdleaderplot_fig

def plot_topSd_pie(basicframe, sd_dateframe, topqty, figlength = 10, figheight = 5, titlesize = 18):
    '''
    Inputs:

    basicframe
    sd_dateframe
    topqty (int): number of top snipers to plot
    kwarg figure size parameters

    Outputs (see docs for frames):

    Pie chart of top sniped
    '''
    #Get Names of the top snipers
    topsds = basicframe.sort_values('Sniped', ascending = False).head(topqty).index.values
    #Make a dataframe without the top snipers
    noleaders_sd_dateframe = sd_dateframe.loc[:, ~sd_dateframe.columns.isin(topsds)]
    #Turn that frame into a single column
    noleaders_sd_dateframe['Sum'] = noleaders_sd_dateframe.sum(axis = 1)
    #Get a frame of just the top snipers
    onlyleaders_sd_dateframe = sd_dateframe.loc[:,topsds]

    sdtopandothers = onlyleaders_sd_dateframe.join(noleaders_sd_dateframe['Sum'])
    sdtopandothers.rename(columns = {'Sum':'Everyone Else'}, inplace = True)
    top5pie_sd_ax = sdtopandothers.sum(axis = 0).plot.pie(figsize = (figlength, figheight), autopct='%1.1f%%')

    top5pie_sd_ax.set_title('Top 5 Sniped', fontsize = titlesize)
    return top5pie_sd_ax



def Generate_Workbook(rawdf, names_dict: dict):
    basicframe = create_basicframe(rawdf, names_dict)
    comboframe = create_comboframe(rawdf, names_dict)
    uniqueframe = create_uniqueframe(comboframe, names_dict)
    daeframe, weekframe, truncdateframe, daterange = create_timetotal(rawdf)
    sn_dateframe, sd_dateframe, sn_byweek, sd_byweek = create_timeind(truncdateframe, names_dict, daterange)
    firstframe = create_firstframe(names_dict, sn_dateframe, sd_dateframe)
    
    awardlist = ['Top 1',
            'Top 2',
            'Top 3',
            'Top sd1',
            'Top sd2',
            'Top sd3',
            'Highest K/D',
            'Most in Day',
            'Most in Day sn',
            'Most in Week',
            'Most in Week sd',
            'Couple',
            'Most Unique',
            'Closest',
            'Farthest',
            'First',
            'Last to be Sniped',
            'Most Sniped Days',
            'Last to get a Snipe',
            'Most Sniping Days',
            'Earliest',
            'OTS',
            'Survivors',
            'Pacifists']

    #Initialize
    awardframe = pd.DataFrame(columns= ['Person','Value'],index=awardlist)

    #Get the top 3 snipers
    top3 = basicframe.sort_values('Snipes',ascending=False).head(n=3)
    awardframe.iloc[:3]=(top3.reset_index().iloc[:,:2])

    #Get the top 3 sniped
    top3sd = basicframe.sort_values('Sniped',ascending=False).head(n=3)
    awardframe.iloc[3:6]=(top3sd.reset_index().iloc[:,[0,2]])

    #Get the best KD
    topKD = basicframe.sort_values('K/D',ascending=False).head(n=1)
    awardframe.iloc[6]=(topKD.reset_index().iloc[:,[0,3]].values)

    #Get most in day and week
    #Initialize empty 2 column array
    arr = np.empty((1,2),dtype = object)

    for frame in [sn_dateframe, sd_dateframe, sn_byweek, sd_byweek]:
        #Get the indicies of the maximum
        date, person = np.where(frame.values == frame.values.max())
        #Append the name/date of the indicies to the awardframe
        arr = np.vstack((arr, [frame.columns[person].values[0], frame.index[date].values[0]]))

    awardframe.iloc[7:11] = arr[1:]

    #Repeat process for the ComboFrame
    sniper, sniped = np.where(comboframe.values == comboframe.values.max())
    awardframe.iloc[11] = [(comboframe.index[sniper].values[0], comboframe.columns[sniped].values[0]), comboframe.values.max()]

    #Use the same process as KD max for unique max
    unique = uniqueframe.sort_values('No. Diff People Sniped',ascending = False).head(n=1)
    awardframe.iloc[12] = unique.reset_index().values

    #Iterate over the FirstFrame columns to get the max from each
    lst = []
    for col in firstframe.columns.values:
        name = firstframe.sort_values(col, ascending = False).head(n=1).reset_index().values
        val = firstframe.sort_values(col, ascending = False).head(n=1).reset_index()[col].values
        lst.append((name,val))
        
    awardframe.loc['Last to get a Snipe'] = lst[0]
    awardframe.loc['Most Sniping Days'] = lst[1]
    awardframe.loc['Last to be Sniped'] = lst[2]
    awardframe.loc['Most Sniped Days'] = lst[3]

    #make a frame of just people that have no snipes
    tempframe = basicframe.loc[basicframe['Snipes'] == 0]
    #Sort out people who havent been sniped
    pacifistframe = tempframe.loc[tempframe['Sniped'] != 0]
    #Add to the award frame
    awardframe.loc['Pacifists'] = [pacifistframe.index.values , 0]

    for i in [69,420,513]:
        awardframe.loc[i] = [[rawdf.iloc[i,0:2].values],rawdf.iloc[i,2]]
        
        
    awardframe.loc['1'] = [[rawdf.iloc[0,0:2].values],rawdf.iloc[0,2]]
    for i in range(99,len(rawdf),100):
        awardframe.loc[i + 1] = [[rawdf.iloc[i,0:2].values],rawdf.iloc[i,2]]

    return awardframe


