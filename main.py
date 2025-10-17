try:
    import sys
    import pandastable as pt
    import tkinter as tk
    from tkinter import ttk, simpledialog
    from PIL import Image, ImageTk
    import os
    import pygsheets
    import json
    from gui.utils import gui_funcs
    from matplotlib import pyplot as plt
    import numpy as np
    import statsmodels.api as sm
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

except:
    print('There was an error importing packages')
    input()
    import sys
    sys.exit()


class mainEntryFrame(ttk.Frame):
    class BookObj:
        #BookObj is essentially jsut an object for keeping track of all the container frames that house the entry frames.
        #It has a tablst which is a list of the container frames, a framelist which is a list of all the subEntryFrames that exist,
        #and activetab, which is the container frame that is currently shown.
        #its only method is buildDataList which returns a list of tuples (sniper,sniped)
        def __init__(self, master):
            container_frame = ttk.Frame(master = master)
            self.tablst = [container_frame]
            self.framelist = []
            self.activetab = 0
        def buildDataList(self):
            self.framedata = []
            for i in self.framelist:
                self.framedata.append((i.sniper_ent.get(),i.sniped_ent.get()))
            return self.framedata
        
    class subEntryFrame:
    #Entry Frame object is one set of Sniper/Sniped entries
    #self.frame is the frame and can be packed/called accordingly
    #self.sniper_ent and self.sniped_ent are the two entry widgets
        def __init__(self,master):
            self.frame = tk.Frame(master=master)

            self.sniper_ent = ttk.Entry(master=self.frame,
                                  width=50,
                                       style = 'entry_normal.TEntry')
            sniper_lbl = ttk.Label(master=self.frame,
                                  text='   Sniper: ',
                                  style = 'default.TLabel') #3 spaces to space from edge
                                  
            self.sniper_ent.grid(row=0, column=1)
            sniper_lbl.grid(row=0, column=0, sticky='nsew')

            self.sniped_ent = ttk.Entry(master=self.frame,
                                  width=50,
                                  style = 'entry_normal.TEntry')
            sniped_lbl = ttk.Label(master=self.frame,
                                  text='  Sniped: ',
                                  style = 'default.TLabel') #3 spaces to space from edge
                                  
            self.sniped_ent.grid(row=1, column=1)
            sniped_lbl.grid(row=1, column=0, sticky = 'nsew')
            
    def __init__(self, master):
        ttk.Frame.__init__(self)
        self.options_frm = ttk.Frame(master = self)
        self.Tabs = self.BookObj(master = self)
        self.fwdbck_frm = ttk.Frame(master = self)
        self.subcan_frm = ttk.Frame(master = self)

        #Options, then fwdbck, then subcan are filled. Book is filled last because I said so

        #place containers
        self.options_frm.grid(row = 0,
                        column = 0,
                             sticky = 'ew')

        self.Tabs.tablst[0].grid(row = 1,
                column = 0)

        #fwdbck_frm not gridded until it is used
        self.subcan_frm.grid(row = 3,
                       column = 0,
                        pady = 7)

        #Options Frame
        #Managed by PACK

        #Whitelist button
        try:
            whitelst_img = Image.open(self.master.location + r'\gui\assets\White_List_PNG.png')
        except Exception as err:
            gui_funcs.GenericError(err)

        whitelst_img= whitelst_img.resize((22,22))
        whitelist_pic= ImageTk.PhotoImage(whitelst_img, master = self.options_frm)

        self.whitelst_btn = ttk.Button(master=self.options_frm,
                                 image=whitelist_pic,
                                 command = master.launchWhiteListWindow,
                                 style = 'default.TButton')
        self.whitelst_btn.image = whitelist_pic

        self.whitelst_btn.pack(side='right')

        #Stats button
        try:
            stats_img = Image.open(self.master.location + r'\gui\assets\Stats_Icon_PNG.png')
        except Exception as err:
            gui_funcs.GenericError(err)

        stats_img= stats_img.resize((22,22))
        stats_pic= ImageTk.PhotoImage(stats_img, master = self.options_frm)

        self.stats_btn = ttk.Button(master=self.options_frm,
                                 image=stats_pic,
                                 command = master.launchStats,
                                 style = 'default.TButton')
        self.stats_btn.image = stats_pic

        self.stats_btn.pack(side='right')

        #QuickView Button
        try:
            qv_img = Image.open(self.master.location + r'\gui\assets\QuickView_Icon_PNG.png')
        except Exception as err:
            gui_funcs.GenericError(err)

        qv_img= qv_img.resize((22,20))
        qv_pic= ImageTk.PhotoImage(qv_img, master = self.options_frm)

        self.qv_btn = ttk.Button(master = self.options_frm,
                                        image = qv_pic,
                                        command = master.launchQuickView,
                                        style = 'default.TButton')
        self.qv_btn.image = qv_pic

        self.qv_btn.pack(side = 'right')

        # Yesterday button
        self.yday_btn = ttk.Button(master = self.options_frm,
                             text = 'Today',
                             command = master.YesterdayBtn,
                             style = 'default.TButton')
        self.yday_btn.pack(side='right')
        
        #Settings Button
        try:
            settings_img = Image.open(self.master.location + r'\gui\assets\Settings_Icon_PNG.png')
        except Exception as err:
            gui_funcs.GenericError(err)

        settings_img= settings_img.resize((22,22))
        settings_pic= ImageTk.PhotoImage(settings_img, master = self.options_frm)

        self.settings_btn = ttk.Button(master = self.options_frm,
                                      image=settings_pic,
                                      command = master.SettingsBtn,
                                      style = 'default.TButton')
        self.settings_btn.image = settings_pic
        self.settings_btn.pack(side = 'left', padx = (2,0))
        

        #Fill in the Foward/Back Frame
        self.fwd_btn = ttk.Button(master=self.fwdbck_frm,
                           text='>',
                           command = master.TabUp,
                            style = 'default.TButton')
        self.fwd_btn.pack(side='right')

        self.bck_btn = ttk.Button(master=self.fwdbck_frm,
                           text='<',
                           command = master.TabDown,
                            style = 'default.TButton')
        self.bck_btn.pack(side='left')

        #Submit/Cancel Frame
        self.sub_btn = ttk.Button(master=self.subcan_frm,
                            text='Submit',
                            command=master.EntrySubmitBtn,
                            style = 'default.TButton')
        self.sub_btn.pack(side='right',padx=10)
        self.can_btn = ttk.Button(master=self.subcan_frm,
                            text='Cancel',
                            command=master.CancelBtn,
                           style = 'default.TButton')
        self.can_btn.pack(side='left',padx=10)

        #Book Creation
        init_frm = self.subEntryFrame(self.Tabs.tablst[-1])
        init_frm.frame.pack()
        self.Tabs.framelist.append(init_frm)
        init_frm.sniped_ent.bind('<Return>', master.EnterKey_handler)

class TableFrame(ttk.Frame):
    def __init__(self):
        ttk.Frame.__init__(self)
        self.IsStuffCreated = False
        
    def updateData(self):
        #Table Window takes the dataframe output of PrepareData and puts it into a table

        #convert list to dataframe
        yday = self.master.mastent_frm.yday_btn.config()['style'][4] == 'yday_green.TButton'
        self.tableframe = gui_funcs.PrepareData(self.master.mastent_frm.Tabs.buildDataList(), yday)
    
    def createFirstTime(self, master):
        #Builds the table frame buttons and table and all that
        self.disp_frm = ttk.Frame(master = self)
        self.dispbtn_frm = ttk.Frame(master = self)

        self.disp_frm.pack()
        
        self.updateData()
        
        #Create and pack the table
        self.table = pt.Table(self.disp_frm,
                         dataframe = self.tableframe,
                         showtoolbar = True)
        self.table.show()

        #Create and pack the buttons
        self.sub_btn = ttk.Button(master = self.dispbtn_frm,
                            text = 'Submit',
                            command = master.DataSubmitBtn,
                            style = 'default.TButton')
        self.sub_btn.pack(side='right',
                     padx=5,
                     pady=10)
        self.can_btn = ttk.Button(master = self.dispbtn_frm,
                             text = 'Cancel',
                            command = master.CancelBtn,
                            style = 'default.TButton')
        self.can_btn.pack(side = 'left',
                     padx = 5,
                     pady = 10)


        self.back_btn = ttk.Button(master = self.dispbtn_frm,
                                   text = 'Back',
                            command = master.BackBtn,
                            style = 'default.TButton')
        self.back_btn.pack(padx = 5,
                     pady = 10)

        self.dispbtn_frm.pack()
        self.IsStuffCreated =  True
    
    def showTable(self, master):
        self.updateData()
        self.table.model.df = self.tableframe
        self.table.redraw()

class WhiteListWindow(tk.Toplevel):
    def __init__(self, master):
        tk.Toplevel.__init__(self)
        self.grab_set()
        
        self.Master = master

        self.mastfrm = ttk.Frame(master = self)
        self.txtent = tk.Text(master = self.mastfrm,
                        width = 40,
                        height = 20)

        self.namefile = self.master.Whitelist
        self.txtent.insert('end','\n'.join(self.namefile))
        self.txtent.pack()

        self.mastbtn_frm = ttk.Frame(master = self.mastfrm)
        self.subcanbtn_frm = tk.Frame(master = self.mastbtn_frm)
        self.filecontrol_frm = tk.Frame(master = self.mastbtn_frm)

        self.sub_btn = ttk.Button(master = self.subcanbtn_frm,
                           text = 'Submit',
                           command = self.WriteNameFile,
                           style = 'whitelist.TButton')
        self.can_btn = ttk.Button(master = self.subcanbtn_frm,
                            text = 'Cancel',
                            command = self.ExitNameWin,
                            style = 'whitelist.TButton')

        self.sub_btn.pack(side='left', padx=1)
        self.can_btn.pack(side='right', padx=1)

        self.subcanbtn_frm.pack(side='left')
        self.filecontrol_frm.pack(side='right')

        self.mastbtn_frm.pack(fill = 'x',
                        pady=3)

        self.mastfrm.pack(padx=10,
                    pady=10)

        self.mainloop()
    
    def WriteNameFile(self):
        newWhiteList = self.txtent.get('1.0','end').split('\n')
        
        sheet = self.master.workbook.worksheet('title','Whitelist')
        #clear and rewrite
        sheet.clear()
        sheet.update_col(1,newWhiteList)
        
        sheet.sort_range((1,1), (len(newWhiteList),1), sortorder='ASCENDING')
        self.master.Whitelist = sheet.get_col(1,include_tailing_empty = False)
        self.destroy()

    def ExitNameWin(self):
        self.destroy()

class SettingsPopUp(tk.Toplevel):
    #Settings PopUp is a developer only tab that allows a person to edit the JSON file with all the program info
    
    def __init__(self, master):
        #Standard Startup Stuff
        tk.Toplevel.__init__(self)
        self.grab_set()
        
        self.Master = master
        
        self.setting_frm = tk.Frame(master = self)
        self.setting_frm.pack()
        
        #Create an empty list to easily get the entered data at the end
        self.entry_list = []
        
        #Iterate through the dictionary that is ProgInfo to display the current info
        for idx, key in enumerate(master.proginfo):
            lbl = ttk.Label(master = self.setting_frm,
                           text = key + ': ')
            ent = ttk.Entry(master = self.setting_frm,
                           width = 80)
            ent.insert('end',master.proginfo[key])
            
            lbl.grid(row = idx, column = 0, sticky = 'e')
            ent.grid(row = idx, column = 1)
            
            #Append the label and entry into a list of tuples
            #These are not lbl['text'] and ent.get() because that info needs to be able to change by the end
            self.entry_list.append((lbl,ent))
            
            
         #Submission and cancel buttons   
        submit_btn = ttk.Button(master = self.setting_frm,
                            text = 'Submit',
                            command = self.SettingSubmitBtn,
                            style = 'default.TButton')
        submit_btn.grid(row = len(master.proginfo),
                    column = 0)
        
        cancel_btn = ttk.Button(master = self.setting_frm,
                            text = 'Cancel',
                            command = self.SettingCancelBtn,
                            style = 'default.TButton')
        cancel_btn.grid(row = len(master.proginfo),
                    column = 1)
    
    #Make sure they really want to do this
    #Interesting quirk, data can be edited while the confirmation window is open and still be succesffully written
    def SettingSubmitBtn(self):
        if tk.messagebox.askyesno(title = 'Submission Confirmation',
                                 message = 'Are you sure you want to edit the Program Info file?\nCheck your edits, this app has very few safety features'):
            templst = []
            
            #Rebuid the ProgInfo File
            for label, entry in self.entry_list:
                templst.append((label['text'][:-2], entry.get()))

            self.master.proginfo = dict(templst)
            
            #Write to ProgInfo
            with open(self.master.location + r'\gui\config\ProgInfo.json','w') as file:
                json.dump(dict(templst), file, indent = 4)

            
            tk.messagebox.showinfo('Nice Job Bucko', 'You have successfully Tampered with the Program File.\nMay God have mercy on your soul')
            self.destroy()
        
        else:
            tk.messagebox.showinfo('No Changes Made', 'Program File Intact\nNo Changes Made')
        
    def SettingCancelBtn(self):
        tk.messagebox.showinfo('No Changes Made', 'Program File Intact\nNo Changes Made')
        self.destroy()

class QuickView(tk.Toplevel):
    def __init__(self, parent):
        from analytics import analytics_funcs as af
        self.parent = parent
        #Fetch data and create necessary 
        raw, names_dict = af.import_data(self.parent.creds, "Sniper Logging", "David", "Whitelist for Decryption")
        basicframe = af.create_basicframe(raw, names_dict)
        dateframe, weekframe, truncdateframe, daterange = af.create_timetotal(raw)

        #Convert snipes per day to total on each day
        sn_daytotal = dateframe.cumsum()

        #Convert date into numerical day
        sn_daytotal = sn_daytotal.reset_index().rename(columns = {'index':'Date'})
        sn_daytotal['Date'] = (sn_daytotal['Date'] - sn_daytotal['Date'].min()).dt.days

        # #Create a least squares fit. No constant because 0 days should equal 0 snipes
        x = sn_daytotal['Date'].values
        y = sn_daytotal['Snipes'].values
        model = sm.OLS(y, x).fit()

        #Set a prediction range and generate predictions using the model
        proj_days = int(self.parent.proginfo["Projection Days"])
        proj = np.arange(0,proj_days+1)
        predictions = model.get_prediction(proj)
        pred_summary = predictions.summary_frame(alpha=0.05) #summary frame has a bunch more information too

        #Now generate the top 3s, this code straight from the award book
        top3_sn_frame = basicframe.sort_values('Snipes',ascending=False).head(n=3).reset_index(names = "Person")
        top3_sd_frame = basicframe.sort_values('Sniped',ascending=False).head(n=3).reset_index(names = "Person")
        #And grab the top KD while were at it
        kdidx = basicframe['K/D'].idxmax()
        kd = basicframe.loc[kdidx]

        #Get the last n00th snipe i.e. 500th, 700th
        idx100 = len(raw) - len(raw)%100 - 1
        last100 = raw.iloc[idx100]

        sn = last100['Sniper']
        sd = last100['Sniped']

        ##########################################################################33
        ### Initialize the Window and start doing all the GUI stuff
        tk.Toplevel.__init__(self)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        

        #Container frame so the theme applies to the background
        container_frm = ttk.Frame(master = self, style = 'default.TFrame')
        container_frm.pack()

        #Master Frames. window managed by GRID
        #Title frame is simple, managed by PACK
        title_frm = ttk.Frame(master = container_frm, style = 'default.TFrame')
        #Top 3 for the stats on the left sidebar, managed by PACK
        top3_frm = ttk.Frame(master = container_frm, style = 'default.TFrame')
        #Houses the graph, managed by PACK
        graph_frm = ttk.Frame(master = container_frm, style = 'default.TFrame')
        #For the projection numbers, total days, and total snipes, managed by GRID
        snapshot_frm = ttk.Frame(master = container_frm, style = 'default.TFrame')

        #Place Masters
        title_frm.grid(row = 0, column = 0, columnspan = 2, pady = 5)
        top3_frm.grid(row = 1, column = 0, rowspan = 2, padx = 5, pady = 5, sticky = 'n')
        graph_frm.grid(row = 1, column = 1)
        snapshot_frm.grid(row = 2, column = 1, pady = (0,5))

        #Title
        title_lbl = ttk.Label(master = title_frm, 
                              text = "StatShot",
                              style = 'title.TLabel')
        title_lbl.pack()

        #Before doing the top3s, set a common indent
        indent = 10
        #All the labels in the top3_frm go XXXX_lbl1 for the title label and XXXX_lbl2 for the actual information
        #Top 3 Snipers
        t3sn_lbl1 = ttk.Label(master = top3_frm, 
                             text = "Top 3 Snipers:",
                             style = 'sub_title.TLabel')        
        t3sn_lbl2 = ttk.Label(master = top3_frm,
                              text = "1: {0}, {1}\n".format(top3_sn_frame["Person"][0], top3_sn_frame["Snipes"][0]) +
                            "2: {0}, {1}\n".format(top3_sn_frame["Person"][1], top3_sn_frame["Snipes"][1]) +
                            "3: {0}, {1}".format(top3_sn_frame["Person"][2], top3_sn_frame["Snipes"][2]),
                              style = 'stat_info.TLabel')
        t3sn_lbl1.pack(pady = (0,5), anchor = 'w')
        t3sn_lbl2.pack(padx = indent, anchor = 'w')

        #Top 3 Sniped
        t3sd_lbl1 = ttk.Label(master = top3_frm, 
                             text = "Top 3 Sniped:",
                             style = 'sub_title.TLabel')                            
        t3sd_lbl2 = ttk.Label(master = top3_frm,
                              text = "1: {0}, {1}\n".format(top3_sd_frame["Person"][0], top3_sd_frame["Sniped"][0]) +
                            "2: {0}, {1}\n".format(top3_sd_frame["Person"][1], top3_sd_frame["Sniped"][1]) +
                            "3: {0}, {1}".format(top3_sd_frame["Person"][2], top3_sd_frame["Sniped"][2]),
                              style = 'stat_info.TLabel')
        t3sd_lbl1.pack(pady = (15,5), anchor = 'w')
        t3sd_lbl2.pack(padx = indent, anchor = 'w')

        #KD
        kd_lbl1 = ttk.Label(master = top3_frm,
                            text = "K/D Leader:",
                            style = 'sub_title.TLabel')
        kd_lbl2 = ttk.Label(master = top3_frm,
                            text = "{0}, {1}".format(kd.name, np.round(kd['K/D'], 2)),
                            style = 'stat_info.TLabel')
        kd_lbl1.pack(pady = (15,5), anchor = 'w')
        kd_lbl2.pack(padx = indent, anchor = 'w')

        #Last n00th
        last100_lbl1 = ttk.Label(master = top3_frm, 
                                text = "{0}th Snipe:".format(idx100+1),
                                style = 'sub_title.TLabel')
        last100_lbl2 = ttk.Label(master = top3_frm,
                                 text = "{0}, {1}".format(sn, sd),
                                 style = 'stat_info.TLabel')
        last100_lbl1.pack(pady = (15,0), anchor = 'w')
        last100_lbl2.pack(padx = indent, anchor = 'w')

        #Graph Frame
        # create the plots from the statsmodel
        fig = plt.figure(1, figsize = (6,3), facecolor= parent.common_background)
        plt.plot(proj, pred_summary['mean'].values, 'k')
        plt.plot(proj, pred_summary['mean_ci_upper'].values, 'r')
        plt.plot(proj, pred_summary['mean_ci_lower'].values, 'r')
        plt.plot(x, y, 'b')

        plt.xlabel("Days", color = self.parent.common_text)
        plt.ylabel("Snipes", color = self.parent.common_text)
        plt.tick_params(axis='x', colors=self.parent.common_text)        
        plt.tick_params(axis='y', colors=self.parent.common_text)
        plt.tight_layout()

        # Create a canvas and place
        canvas = FigureCanvasTkAgg(fig,
                                master = graph_frm)  
        canvas.draw()
        canvas.get_tk_widget().pack()

        #Snapshot frame
        mean = int(np.round(pred_summary['mean'][proj_days], 0))
        upper = int(np.round(pred_summary['mean_ci_upper'][proj_days], 0))
        lower = int(np.round(pred_summary['mean_ci_lower'][proj_days], 0))

        mean_lbl = ttk.Label(master = snapshot_frm,
                            text = 'Projection: ' + str(mean),
                            style = 'projection_info.TLabel')
        upper_lbl = ttk.Label(master = snapshot_frm,
                             text = 'Upper Projection: ' + str(upper),
                            style = 'projection_info.TLabel')
        lower_lbl = ttk.Label(master = snapshot_frm,
                             text = 'Lower Projection: ' + str(lower),
                            style = 'projection_info.TLabel')

        mean_lbl.grid(row = 0, column = 2, padx = 10, columnspan = 2)
        upper_lbl.grid(row = 0, column = 4, sticky = 'e', columnspan = 2)
        lower_lbl.grid(row = 0, column = 0, sticky = 'w', columnspan = 2)

        totalday_lbl = ttk.Label(master = snapshot_frm,
                                 text = "Total Days: {}".format(len(sn_daytotal['Date'])),
                                 style = 'projection_info.TLabel')
        totalsn_lbl = ttk.Label(master = snapshot_frm,
                                text = "Total Snipes: {}".format(len(raw)-1),
                                style = 'projection_info.TLabel')

        totalday_lbl.grid(row = 1, column = 2, padx = 5, pady = 5)
        totalsn_lbl.grid(row = 1, column = 3, padx = 5)

    def on_close(self):
        plt.close('all')
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        #Self.location is the location where the folder is stored
        self.location = os.path.dirname(__file__)
        try:
            with open(self.location + r'\gui\config\ProgInfo.json', 'r') as progfile:
             self.proginfo = json.load(progfile)
        except Exception as err:
            gui_funcs.GenericError(err)
            sys.exit()
        
        if self.proginfo['Google Sheet Name'] == '':
            self.proginfo['Google Sheet Name'] = simpledialog.askstring('No Google Sheet', 'No Google Sheet Found in Program Info File\nEnter Google Sheet Name:')
            with open(self.location + r'\gui\config\ProgInfo.json','w') as file:
                json.dump(dict(self.proginfo), file, indent = 4)
        
        #Load The Whitelist    
        #Call the Sheet
        exit = False
        while exit == False:
            try:
                self.creds = gui_funcs.oauth_flow(self.location)
                self.client = pygsheets.authorize(credentials = self.creds)
                self.workbook = self.client.open(self.proginfo['Google Sheet Name'])
                #Get the whole column of the whitelist sheet
                self.Whitelist = self.workbook.worksheet('title','Whitelist').get_col(1,include_tailing_empty = False)
                exit = True
            except pygsheets.exceptions.SpreadsheetNotFound:
                self.proginfo['Google Sheet Name'] = simpledialog.askstring('Sheet Not Found', 'Sheet Name \"{}\" Not Found\nPlease Re-enter:'.format(self.proginfo['Google Sheet Name']))
                if self.proginfo['Google Sheet Name'] == None:
                    tk.messagebox.showerror('No Sheet Entered', 'No Google Sheet Name Entered\nExiting Program')
                    sys.exit()
                with open(self.location + r'\gui\config\ProgInfo.json','w') as file:
                    json.dump(dict(self.proginfo), file, indent = 4)
            except Exception as e:
                tk.messagebox.showerror('Google Error', f'A Google Error occurred on startup\n\n{e}')
                sys.exit()

        if self.proginfo['User'] == '':
            exit = False
            while exit == False:
                self.proginfo['User'] = simpledialog.askstring('Bad User','Invalid User Name \"{}\"\nEnter User Name:'.format(self.proginfo['User']))
                if self.proginfo['User'] == None:
                    tk.messagebox.showerror('No User Entered', 'No User Name Entered\nExiting Program')
                    sys.exit()
                try:
                    self.workbook.worksheet('title',self.proginfo['User'])
                    with open(self.location + r'\gui\config\ProgInfo.json','w') as file:
                        json.dump(dict(self.proginfo), file, indent = 4)
                    exit = True
                except pygsheets.exceptions.WorksheetNotFound:
                    continue
                except Exception as e:
                    gui_funcs.GenericError(e)
                    sys.exit()
                
        #Initialize
        # ThemedTk.__init__(self, theme = 'clam')
        tk.Tk.__init__(self)
        self.style = ttk.Style()
        self.style.layout('TNotebook.Tab', [])
        
        #Set Title and Icon
        try:
            self.iconbitmap(self.location + r'\gui\assets\CrossHairIcon_ICO.ico')
        except Exception as err:
            gui_funcs.GenericError(err)
        
        self.title('Data Entry')

        #Style Block
        #Main Window
        self.style.configure('default.TButton', font=('Helvetica', 12))
        self.style.configure('entry_error.TEntry', fieldbackground='red')
        self.style.configure('entry_normal.TEntry', fieldbackground='white')
        self.style.configure('default.TLabel', font = ('Segoe UI', 12))
        self.style.configure('yday_green.TButton', font=('Segoe UI', 11, 'bold'), foreground = '#339E37')

        #Whitelist Window
        self.style.configure('whitelist.TButton', width = 7, font = ('Segoe UI', 10))
        
        #QuickView Window
        self.common_background = "#404145"
        self.common_text = "#FFFFFF"
        self.style.configure('title.TLabel', font = ('Bahnschrift', 24, 'bold'), background = self.common_background, foreground = self.common_text)
        self.style.configure('sub_title.TLabel', font = ('Segoe UI', 12, 'bold'), background = self.common_background, foreground = self.common_text)
        self.style.configure('stat_info.TLabel', font = ('Gill Sans MT', 10), background = self.common_background, foreground = self.common_text)
        self.style.configure('projection_info.TLabel', font = ('Gill Sans MT', 10, 'italic'), background = self.common_background, foreground = self.common_text)
        self.style.configure('default.TFrame', background = self.common_background)
        
        #Start both master frames and pack the entry frame        
        self.mastent_frm = mainEntryFrame(self)
        self.masttable_frm = TableFrame()
        self.mastent_frm.pack()
        
        
    def launchWhiteListWindow(app):
        WhiteListWindow(app)

    def launchQuickView(app):
        QuickView(app)
    
    def YesterdayBtn(app):
        #Check what style it is, and change it to the other
        if app.mastent_frm.yday_btn.config()['style'][4] == 'default.TButton':
            app.mastent_frm.yday_btn.config(text = 'Yesterday', style = 'yday_green.TButton')
        elif app.mastent_frm.yday_btn.config()['style'][4] == 'yday_green.TButton':
            app.mastent_frm.yday_btn.config(text = 'Today', style = 'default.TButton')
        else:
            raise Exception('Error in Yesterday Button')
        
    def EntrySubmitBtn(app):
        #Make Sure the data is good before sending it on
        
        if gui_funcs.CheckWhiteList(app.mastent_frm.Tabs.framelist,app):
            if app.masttable_frm.IsStuffCreated:
                app.mastent_frm.pack_forget()
                app.masttable_frm.showTable(app)
                app.masttable_frm.pack()
            else:
                app.mastent_frm.pack_forget()
                app.masttable_frm.createFirstTime(app)
                app.masttable_frm.pack()
            
        else:
            tk.messagebox.showinfo('Bad Value', 'Illegal Name Entered\nPlease Re-enter')
            
            
    def CancelBtn(app):
        tk.messagebox.showinfo('Program Finish','Program Ended\nNo Data Written\n\nHave a Nice Day!!!')
        app.destroy()
        sys.exit()
        
    def DataSubmitBtn(app):
        #Open WorkBook
        try:
            #Access the worksheet
            worksheet = app.workbook.worksheet('title', app.proginfo['User'])
            #Append table returns sheetinfo, which has a couple fun tidbits about the worksheet
            sheetinfo = worksheet.append_table(app.masttable_frm.table.model.df.values.tolist())

            #num is the index of the final cell in the range
            num = sheetinfo['updates']['updatedRange'].end[0]

            #get the last 100 and pull out the sniper/sniped cells
            last100idx = num - num%100
            sncell = worksheet.cell((last100idx + 1, 1))
            sdcell = worksheet.cell((last100idx + 1, 2))

            tk.messagebox.showinfo('Program Success',('Data Successfully Written\n{0} Snipes Recorded\n\n{1}th Snipe: {2}, {3}'.format(num, last100idx, sncell.value, sdcell.value)))
            app.destroy()
            sys.exit()

        except Exception as err:
            #If things go catastrophically wrong, the data needs to be saved
            data = app.masttable_frm.table.model.df
            location = os.path.dirname(app.location) + r'\\Data_Backup.xlsx'
            data.to_excel(location )

            tk.messagebox.showerror('Write Error','There was an error while writing to the Google Sheet\n\n{0}\n\n' \
            'Your data has been saved to the following file:\n\n' \
            '{1}\n\nProgram will now close'.format(err,location))

            app.destroy()
            sys.exit()

    def BackBtn(app):
        app.masttable_frm.pack_forget()
        app.mastent_frm.pack()
        
    def SettingsBtn(app):
        SettingsPopUp(app)

    def launchStats(app):
        if tk.messagebox.askyesno(title = 'Launch Analytics',
                                 message = 'This will close the current window and open the Analytics Window\nAll current entries will be lost\n\nContinue?'):
            from subprocess import Popen
            Popen([sys.executable, app.location + r'\analytics_gui.py'])
            print('Launched')
            app.destroy()

        else:
            return
        
    def TabUp(app):
        #Check within legal range
        if app.mastent_frm.Tabs.activetab >= len(app.mastent_frm.Tabs.tablst) - 1:
            pass
        else:
            #Hide current, reset active, show active
            app.mastent_frm.Tabs.tablst[app.mastent_frm.Tabs.activetab].grid_forget()
            app.mastent_frm.Tabs.activetab+=1
            app.mastent_frm.Tabs.tablst[app.mastent_frm.Tabs.activetab].grid(row = 1, column = 0)

    def TabDown(app):
        #Check within legal range
        if app.mastent_frm.Tabs.activetab <= 0:
            pass
        else:
            #Hide current, reset active, show active
            app.mastent_frm.Tabs.tablst[app.mastent_frm.Tabs.activetab].grid_forget()
            app.mastent_frm.Tabs.activetab-=1
            app.mastent_frm.Tabs.tablst[app.mastent_frm.Tabs.activetab].grid(row = 1, column = 0)
        
    def EnterKey_handler(app, event):
        if len(app.mastent_frm.Tabs.framelist) % int(app.proginfo['Entries per Tab']) == 0:
            #Check the whitelist. First arg must be a list, second must be the app
            if gui_funcs.CheckWhiteList([app.mastent_frm.Tabs.framelist[-1]],app):
                app.mastent_frm.Tabs.framelist[-1].sniped_ent.unbind('<Return>')
                
                #Check if this is the first time a new tab is being created
                if len(app.mastent_frm.Tabs.framelist) == int(app.proginfo['Entries per Tab']):
                    #Grid the Forward/Back Buttons
                    app.mastent_frm.fwdbck_frm.grid(row = 2,
                                               column = 0)
                    
                #Create another container frame
                container_frm = ttk.Frame(master = app.mastent_frm)
                #Remove the current container frame
                app.mastent_frm.Tabs.tablst[app.mastent_frm.Tabs.activetab].grid_forget()
                #Show the new one
                container_frm.grid(row = 1, column = 0, padx = 5)
                #Add new container to the tablst and make it the active frame
                app.mastent_frm.Tabs.tablst.append(container_frm)
                app.mastent_frm.Tabs.activetab = len(app.mastent_frm.Tabs.tablst) - 1
                
                #Create and pack the first subEntryFrame
                Frame = app.mastent_frm.subEntryFrame(container_frm)
                Frame.frame.pack()
                
                #Add newest frame to the framelist
                app.mastent_frm.Tabs.framelist.append(Frame)
                #Bind the EnterKey_handler to the newest frame
                Frame.sniped_ent.bind('<Return>', app.EnterKey_handler)

            else:
                tk.messagebox.showinfo('Bad Value', 'Illegal Name Entered\nPlease Re-enter')
                
        else:
            #Check the whitelist. First arg must be a list, second must be the app
            if gui_funcs.CheckWhiteList([app.mastent_frm.Tabs.framelist[-1]],app):
                app.mastent_frm.Tabs.framelist[-1].sniped_ent.unbind('<Return>')
                #Create a new frame
                Frame = app.mastent_frm.subEntryFrame(app.mastent_frm.Tabs.tablst[-1])
                Frame.frame.pack()
                #Add newest frame to the framelist
                app.mastent_frm.Tabs.framelist.append(Frame)
                #Bind the EnterKey_handler to the newest frame
                Frame.sniped_ent.bind('<Return>', app.EnterKey_handler)

            else:
                tk.messagebox.showinfo('Bad Value', 'Illegal Name Entered\nPlease Re-enter')


if __name__ == "__main__":
    app = App()
    app.mainloop()