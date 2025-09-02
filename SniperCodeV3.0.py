import numpy as np
import datetime
import pandas as pd
import openpyxl as pyxl
import sys
import pandastable as pt
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
from PIL import Image, ImageTk
import os
import pygsheets
import json
from gui.utils import gui_funcs

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
                                  text='Sniper: ')
                                  
            self.sniper_ent.grid(row=0, column=1)
            sniper_lbl.grid(row=0, column=0)

            self.sniped_ent = ttk.Entry(master=self.frame,
                                  width=50,
                                  style = 'entry_normal.TEntry')
            sniped_lbl = ttk.Label(master=self.frame,
                                  text='Sniped: ')
                                  
            self.sniped_ent.grid(row=1, column=1)
            sniped_lbl.grid(row=1, column=0)
            
    def __init__(self, master):
        ttk.Frame.__init__(self)
        self.options_frm = ttk.Frame(master = self)
        self.Tabs = self.BookObj(master = self)
        self.fwdbck_frm = ttk.Frame(master = self)
        self.subcan_frm = ttk.Frame(master = self)

        #Below, options, then fwdbck, then subcan are filled. Book is filled last because I said so

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

        #Whitelist button
        try:
            whitelst_img = Image.open(self.master.location + r'\gui\assets\White_List_PNG.png')
        except Exception as err:
            master.fun.GenericError(err)

        whitelst_img= whitelst_img.resize((22,22))
        whitelist_pic= ImageTk.PhotoImage(whitelst_img, master = self.options_frm)

        self.whitelst_btn = ttk.Button(master=self.options_frm,
                                 image=whitelist_pic,
                                 command = master.launchWhiteListWindow,
                                 style = 'default.TButton')
        self.whitelst_btn.image = whitelist_pic

        self.whitelst_btn.pack(side='right')

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
            master.fun.GenericError(err)

        settings_img= settings_img.resize((22,22))
        settings_pic= ImageTk.PhotoImage(settings_img, master = self.options_frm)

        self.settings_btn = ttk.Label(master = self.options_frm,
                                      image=settings_pic,
                                      style = 'default.TLabel')
        self.settings_btn.bind('<Button-1>', master.SettingsBtn)
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
        self.tableframe = self.master.fun.PrepareData(self.master.mastent_frm.Tabs.buildDataList(), yday)
    
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
            with open(r'gui\config\ProgInfo.json','w') as file:
                json.dump(dict(templst), file, indent = 4)

            
            tk.messagebox.showinfo('Nice Job Bucko', 'You have successfully Tampered with the Program File.\nMay God have mercy on your soul')
            self.destroy()
        
        else:
            tk.messagebox.showinfo('No Changes Made', 'Program File Intact\nNo Changes Made')
        
    def SettingCancelBtn(self):
        tk.messagebox.showinfo('No Changes Made', 'Program File Intact\nNo Changes Made')
        self.destroy()
    
class App(ThemedTk):
    def __init__(self):
        #Self.fun is just a way to store all the general purpose functions to keep it more organized
        self.fun = gui_funcs

        #Self.location is the location where the folder is stored
        self.location = os.path.dirname(__file__)
        try:
            with open(self.location + r'\gui\config\ProgInfo.json', 'r') as progfile:
             self.proginfo = json.load(progfile)
        except Exception as err:
            print(err)
            self.fun.GenericError(err)
            sys.exit()
        
        #Load The Whitelist    
        #Call the Sheet
        try:
            creds = gui_funcs.oauth_flow(self.location)
            self.client = pygsheets.authorize(credentials = creds)
            self.workbook = self.client.open(self.proginfo['Google Sheet Name'])
            #Get the whole column of the whitelist sheet
            self.Whitelist = self.workbook.worksheet('title','Whitelist').get_col(1,include_tailing_empty = False)
        except Exception as e:
            tk.messagebox.showerror('Google Error', f'A Google Error occurred on startup\n\n{e}')
            sys.exit()
            
        #Initialize
        ThemedTk.__init__(self, theme = self.proginfo['Theme'])
        self.style = ttk.Style()
        self.style.layout('TNotebook.Tab', [])
        
        #Set Title and Icon
        try:
            self.iconbitmap(self.location + r'\gui\assets\CrossHairIcon_ICO.ico')
        except Exception as err:
            self.fun.GenericError(err)
        
        self.title('Sniper Data Entry V3.0')

        #set button/label styles
        self.style.configure('default.TButton', font=('Helvetica', 12))
        self.style.configure('entry_error.TEntry', fieldbackground='red')
        self.style.configure('entry_normal.TEntry', fieldbackground='white')
        self.style.configure('yday_green.TButton', font=('Helvetica', 11, 'bold'), foreground = '#339E37')
        self.style.configure('whitelist.TButton', width = 7, font = ('Helvetica', 10))
        self.style.configure('hidden.TLabel', width = 5,
                             fieldbackground = self.style.lookup('TFrame','background'),
                             foreground = self.style.lookup('TFrame', 'background'))
        
        #Start both master frames and pack the entry frame
        self.mastent_frm = mainEntryFrame(self)
        self.masttable_frm = TableFrame()
        self.mastent_frm.pack()
        
        
    def launchWhiteListWindow(app):
        WhiteListWindow(app)
    
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
        
        if app.fun.CheckWhiteList(app.mastent_frm.Tabs.framelist,app):
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
        #Open Sheet
        # gc = pygsheets.authorize(service_file=app.proginfo['Credentials'])
        
        #Open WorkBook
        # workbook = gc.open(app.proginfo['Google Sheet Name'])
        worksheet = app.workbook.worksheet('title', app.proginfo['User'])

        sheetinfo = worksheet.append_table(app.masttable_frm.table.model.df.values.tolist())

        num = sheetinfo['updates']['updatedRange'].end[0]

        tk.messagebox.showinfo('Program Success',('Data Successfully Written\n' + str(num-1) + ' Snipes Recorded'))
        app.destroy()
        sys.exit()

    def BackBtn(app):
        app.masttable_frm.pack_forget()
        app.mastent_frm.pack()
        
    def SettingsBtn(app, event):
        SettingsPopUp(app)

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
            if app.fun.CheckWhiteList([app.mastent_frm.Tabs.framelist[-1]],app):
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
                container_frm.grid(row = 1, column = 0)
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
            if app.fun.CheckWhiteList([app.mastent_frm.Tabs.framelist[-1]],app):
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

app = App()
app.mainloop()