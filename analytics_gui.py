import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import filedialog, simpledialog
from analytics.analytics_funcs import *
from gui.utils import gui_funcs
import json
import pandas as pd
import sys
import os


class CtrlPanel(ThemedTk):
    def __init__(self):
        #Load in proginfo
        self.location = os.path.dirname(__file__)
        try:
            with open(self.location + r'\gui\config\ProgInfo.json', 'r') as progfile:
             self.proginfo = json.load(progfile)
        except Exception as err:
            gui_funcs.GenericError(err)
            sys.exit()
        
        #Set the out directory. If it is empty, no big deal that's handled later
        self.out_dir = self.proginfo['Analytics Location']

        #Create dictionary of all the spreadsheet stats
        self.spreadsheet_dict = {"basicframe" : ['Standard Stats'],
                         "comboframe" : ['Combinations'],
                         "uniqueframe" : ["No. Unique People Sniped"],
                         "firstframe" : ["First Snipe/Sniped Dates"],
                         "dateframe" : ["Total Snipes per Day"],
                         "weekframe" : ["Total Snipes per Week"],
                         "sn_dateframe" : ["Ind. Snipes per Day"],
                         "sd_dateframe": ["Ind. Sniped per Day"],
                         "sn_byweek" : ["Ind. Snipes per Week"],
                         "sd_byweek" : ["Ind. Sniped per Week"],
                         }
        
        #Create a second dictionary of all the graph stats
        self.graph_dict = {"weekdaypie" : ['Weekday Pie Chart'],
                           "overtime" : ['Snipes Over Time'],
                      "topSn_overtime" : ['Top Snipers Over Time'],
                      "topSn_pie" : ['Top Snipers Pie Chart'],
                      "topSd_overtime" : ['Top Sniped Over Time'],
                      "topSd_pie" : ['Top Sniped Pie Chart'],
                      }
        ThemedTk.__init__(self)
        self.title("Analytics Control Panel")

        #Create the header frame
        header_frm = tk.Frame(master = self)
        header_frm.pack(padx = 5, pady = (5,0))
        header_toplbl = tk.Label(master = header_frm, text = "Analytics Control Panel", font = ('Arial', 20, 'bold'))
        header_toplbl.pack()
        
        #Create static and dynamic labels
        preset_frm = tk.Frame(master = self)
        preset_frm.pack(pady = (0,5))
        presetSt_lbl = tk.Label(master = preset_frm, text = "Select stats or", font = ('Arial', 10))
        presetDy_lbl = tk.Label(master = preset_frm, text = "select a preset", font = ('Arial', 10, 'underline'), cursor = 'hand2')
        presetDy_lbl.bind("<Button-1>", self.set_preset)

        presetSt_lbl.pack(side = 'left')
        presetDy_lbl.pack(side = 'right')

        #Container for spreadsheet stat options
        self.sheets_frm = tk.Frame(master = self, relief = 'sunken', borderwidth = 2)
        self.sheets_frm.pack(padx = 5, pady = 5, expand=True, fill = 'x')

        #Create all the labels and checkboxes from the dictionary
        for idx, key in enumerate(self.spreadsheet_dict):
            lbl = ttk.Label(master = self.sheets_frm, text = self.spreadsheet_dict[key][0] + " :")
            #The bool_var controls the state of the checkbox
            bool_var = tk.BooleanVar()
            box = ttk.Checkbutton(self.sheets_frm, variable=bool_var)

            #The second value of spreadsheet_dict[key] is the bool_var, not a boolean value
            self.spreadsheet_dict[key].append(bool_var)

            lbl.grid(row = idx, column= 0, sticky = 'w', padx = (0,150), pady = 5)
            box.grid(row = idx, column= 1, sticky = 'e')

        #Container for graph stat options
        self.graph_frm = tk.Frame(master = self, relief = 'sunken', borderwidth = 2)
        self.graph_frm.pack(padx = 5, pady = 5, expand = True, fill = 'x')

        #Create all the labels and checkboxes from the dictionary
        for idx, key in enumerate(self.graph_dict):
            lbl = ttk.Label(master = self.graph_frm, text = self.graph_dict[key][0] + " :")
            #The bool_var controls the state of the checkbox
            bool_var = tk.BooleanVar()
            box = ttk.Checkbutton(self.graph_frm, variable=bool_var)

            #The second value of graph_dict[key] is the bool_var, not a boolean value
            self.graph_dict[key].append(bool_var)

            lbl.grid(row = idx, column= 0, sticky = 'w', padx = (0,170), pady = 5)
            box.grid(row = idx, column= 1, sticky = 'e')

        #Create a subframe to put in the top values
        topnum_frm = tk.Frame(master = self.graph_frm, relief = 'ridge', borderwidth = 2)
        topnum_lbl = tk.Label(master = topnum_frm, text = "Top Qty for Top Sniper/Sniped Graphs: ")
        self.topnum_entry = tk.Entry(master = topnum_frm, width = 2)
        topnum_lbl.pack(side = 'left', padx = (5,0), pady = 5)
        self.topnum_entry.pack(side = 'right', padx = (0,5), pady = 5)
        self.topnum_entry.insert(0, '5')
        topnum_frm.grid(row = len(self.graph_dict), column = 0, columnspan = 2, padx = 2, pady = 10)

        #Create and pack the sbumission and save preset buttons
        self.submit_frm = ttk.Frame(master = self)
        self.submit_frm.pack()

        self.sub_btn = ttk.Button(master = self.submit_frm, text = "Generate", command = self.btncmd)
        self.savePreset_btn = ttk.Button(master = self.submit_frm, text = 'Save as Preset', command = self.save_preset)
        self.sub_btn.pack(side = 'right')
        self.savePreset_btn.pack(side = 'left')

        self.settings_frm = ttk.Frame(master = self, borderwidth = 2, relief = 'sunken')
        self.settings_frm.pack(padx = 4, pady = 4)

        out_lbl = ttk.Label(master = self.settings_frm, text = "Output File Location:")
        out_lbl.grid(row = 0, column = 0, sticky = 'w')

        self.file_lbl = ttk.Label(master = self.settings_frm, text = self.out_dir, width = 40, background = 'darkgray')
        self.file_lbl.grid(row = 1, column = 0, columnspan = 5)
        change_btn = ttk.Button(master = self.settings_frm, text = "Set", command = self.changefile)
        change_btn.grid(row = 1, column = 6)

    def set_preset(self, event):
        
        def set_preset_vals():
            #This is a subfunction because the button needs a function, but passing variables was too complicated
            if box.get() == '':
                tk.messagebox.showerror('No Preset Selected', 'Please select a preset')
                return
            else:
                #Get all the booleans of the preset file
                with open(self.location + r'\analytics\presets\{}'.format(box.get()), 'r') as f:
                    presets = json.load(f)

            #Apply the presets to the bool_vars, which automatically updates the checkboxes
            for key in presets['sheets']:
                self.spreadsheet_dict[key][1].set(presets['sheets'][key])
            for key in presets['graphs']:
                self.graph_dict[key][1].set(presets['graphs'][key])
            top.destroy()
            return
        
        #Start the window and stop interaction with the main window
        top = tk.Toplevel(self)
        top.title('Select Preset')
        top.grab_set()

        lbl = tk.Label(master = top, text = 'Select a Preset: ')
        lbl.pack(side = 'left')

        #The options for the box are the files in \presets. It says .json. I'm too lazy to remove it
        opt_list = os.listdir(self.location + r'\analytics\presets')
        box = ttk.Combobox(master = top, values = opt_list)
        box.pack(side = 'left')

        btn = tk.Button(master = top, text = 'Set', command = set_preset_vals)
        btn.pack(side = 'bottom')

    def btncmd(self):
        #Make sure we have a path
        if not os.path.exists(self.out_dir):
            tk.messagebox.showerror('Bad Path', 'Output location does not exist\nPlease set a new location')
            return

        #Make sure they selected stats
        #If not any(boolvar.get() for *iterate through bool_vars*)
        if not any(self.spreadsheet_dict[key][1].get() for key in self.spreadsheet_dict) and not any(self.graph_dict[key][1].get() for key in self.graph_dict):
            tk.messagebox.showerror('No Options', 'No options selected\nPlease select some stats')
            return

        #Create a little top window to show that hte app is working
        top = tk.Toplevel(self)
        top.title('Working...')
        top.grab_set()

        #StringVar because it's going to get edited a lot
        msgstr = tk.StringVar()
        msgstr.set('Contacting Google...')
        msg = ttk.Label(master = top, textvariable = msgstr)
        msg.pack(padx = 10, pady = 10)

        #Without this it flashes all the updates at the end
        top.update()
        
        #Load The Whitelist    
        #Call the Sheet
        try:
            self.creds = gui_funcs.oauth_flow(self.location)
        except Exception as e:
            tk.messagebox.showerror('Google Error', f'A Google Error occurred on startup\n\n{e}')
            sys.exit()

        msgstr.set(msgstr.get() + '\nFetching Data...')
        top.update()
        #Get data
        try:
            raw, names_dict = import_data(self.creds, self.proginfo['Google Sheet Name'], self.proginfo['User'], self.proginfo['Whitelist Sheet'])
        except pygsheets.exceptions.WorksheetNotFound as e:
            tk.messagebox.showerror('oof')
        except Exception as e:
            gui_funcs.GenericError(e)
            sys.exit()

        #Initialize. Graph flag is used at the end
        output_dic = {}
        graph_flag = False

        msgstr.set(msgstr.get() + '\nProcessing Data...')
        top.update()

        #now run through each option one by one. It has to be this way because the functions have required dataframes
        #Doing it this way means that we don't redundantly generate frames
        try:
            '''Standard:
            if *option selected*:
                try: *Name of required frames*
                except NameError: *generate requirements*
                create frame and append to output_dic
            '''
            if self.spreadsheet_dict['basicframe'][1].get():
                basicframe = create_basicframe(raw,names_dict)
                output_dic['basicframe'] = [basicframe, "Basic"]

            if self.spreadsheet_dict['comboframe'][1].get():
                comboframe = create_comboframe(raw, names_dict)
                output_dic['comboframe'] = [comboframe, "Combo"]

            if self.spreadsheet_dict['uniqueframe'][1].get():
                #Standard block. Check to see if frame exists, if it doesn't, make it
                try:
                    comboframe
                except NameError:
                    comboframe = create_comboframe(raw, names_dict)

                output_dic['uniqueframe'] = [create_uniqueframe(comboframe, names_dict), "Unique"]
            if self.spreadsheet_dict['dateframe'][1].get() or self.spreadsheet_dict['weekframe'][1].get():
                dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)
                if self.spreadsheet_dict['dateframe'][1].get():
                    output_dic['dateframe'] = [dateframe, "Total by Day"]
                    
                if self.spreadsheet_dict['weekframe'][1].get():
                    output_dic['weekframe'] = [weekframe, "Total by Week"]

            if self.spreadsheet_dict['sn_dateframe'][1].get() or self.spreadsheet_dict['sd_dateframe'][1].get() or self.spreadsheet_dict['sn_byweek'][1].get() or self.spreadsheet_dict['sd_byweek'][1].get():
                #All 4 are nested because all their requirements stem from the same function
                try:
                    truncdateframe
                except NameError:
                    dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)

                sn_dateframe, sd_dateframe, sn_byweek, sd_byweek = create_timeind(truncdateframe, names_dict, daterange)

                if self.spreadsheet_dict['sn_dateframe'][1].get():
                        output_dic['sn_dateframe'] = [sn_dateframe, "Ind. Sn. Day"]

                if self.spreadsheet_dict['sd_dateframe'][1].get():
                        output_dic['sd_dateframe'] = [sd_dateframe, "Ind. Sd. Day"]

                if self.spreadsheet_dict['sn_byweek'][1].get():
                        output_dic['sn_byweek'] = [sn_byweek, "Ind. Sn. Week"]

                if self.spreadsheet_dict['sd_byweek'][1].get():
                        output_dic['sd_byweek'] = [sd_byweek, "Ind. Sd. Week"]

            #Check if any graphs are selected
            if any([self.graph_dict[key][1].get() for key in self.graph_dict]):
                msgstr.set(msgstr.get() + '\nDrawing Graphs...')
                top.update()

                #Set a second folder to dump the graphs into
                graph_flag = True
                graph_dir = self.out_dir + r'\Graphs {}'.format(pd.Timestamp.now().strftime('%Y-%m-%d %H-%M'))
                os.mkdir(graph_dir)


                if self.graph_dict['weekdaypie'][1].get():
                    try:
                        truncdateframe
                    except NameError:
                        dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)

                    pie = plot_weekdaypie(truncdateframe)
                    pie.figure.savefig(graph_dir + r'\Weekday Pie Chart.png')

                if self.graph_dict['overtime'][1].get():
                    try:
                        dateframe
                    except NameError:
                        dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)
                    plot = plot_overtime(dateframe)
                    plot.figure.savefig(graph_dir + r'\Snipes Over Time.png')
                
                if self.graph_dict['topSn_overtime'][1].get():
                    try:
                        basicframe
                    except NameError:
                        basicframe = create_basicframe(raw, names_dict)

                    try:
                        sn_dateframe
                    except NameError:
                        #Nested requirements check. Now this is some real advanced stuff
                        try:
                            truncdateframe
                            daterange
                        except NameError:
                            dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)
                        
                        sn_dateframe, sd_dateframe, sn_byweek, sd_byweek = create_timeind(truncdateframe, names_dict, daterange)
                    
                    qty = self.topnum_entry.get()
                    plot = plot_topSn_overtime(basicframe, sn_dateframe, int(qty))
                    plot.savefig(graph_dir + r'\Top Snipers Plot.png')

                        
                if self.graph_dict['topSn_pie'][1].get():
                    try:
                        basicframe
                    except NameError:
                        basicframe = create_basicframe(raw, names_dict)

                    try:
                        sn_dateframe
                    except NameError:
                        #Nested requirements check. Now this is some real advanced stuff
                        try:
                            truncdateframe
                            daterange
                        except NameError:
                            dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)

                        sn_dateframe, sd_dateframe, sn_byweek, sd_byweek = create_timeind(truncdateframe, names_dict, daterange)

                    qty = self.topnum_entry.get()
                    plot = plot_topSn_pie(basicframe, sn_dateframe, int(qty))
                    plot.figure.savefig(graph_dir + r'\Top Snipers Pie Chart')
                    
                if self.graph_dict['topSd_overtime'][1].get():
                    print('sd time')
                    try:
                        basicframe
                    except NameError:
                        basicframe = create_basicframe(raw, names_dict)

                    try:
                        sd_dateframe
                    except NameError:
                        #Nested requirements check. Now this is some real advanced stuff
                        try:
                            truncdateframe
                            daterange
                        except NameError:
                            dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)
                        
                        sn_dateframe, sd_dateframe, sn_byweek, sd_byweek = create_timeind(truncdateframe, names_dict, daterange)
                    
                    qty = self.topnum_entry.get()
                    plot = plot_topSd_overtime(basicframe, sd_dateframe, int(qty))
                    plot.savefig(graph_dir + r'\Top Sniped Plot.png')

                        
                if self.graph_dict['topSd_pie'][1].get():
                    print('sd pie')
                    try:
                        basicframe
                    except NameError:
                        basicframe = create_basicframe(raw, names_dict)

                    try:
                        sd_dateframe
                    except NameError:
                        #Nested requirements check. Now this is some real advanced stuff
                        try:
                            truncdateframe
                            daterange
                        except NameError:
                            dateframe, weekframe, truncdateframe, daterange = create_timetotal(raw)

                        sn_dateframe, sd_dateframe, sn_byweek, sd_byweek = create_timeind(truncdateframe, names_dict, daterange)

                    qty = self.topnum_entry.get()
                    plot = plot_topSn_pie(basicframe, sd_dateframe, int(qty))
                    plot.figure.savefig(graph_dir + r'\Top Sniped Pie Chart')
    
        except Exception as e:
            gui_funcs.GenericError(e)
            top.destroy()
            return
        
        msgstr.set(msgstr.get() + '\nWriting File...')
        top.update()

        if len(output_dic) > 0:
            out_file = self.out_dir + r'\SnipeStats {}.xlsx'.format(pd.Timestamp.now().strftime('%Y-%m-%d %H-%M'))
            writer = pd.ExcelWriter(out_file, engine='openpyxl')
            for key in output_dic:
                output_dic[key][0].to_excel(writer, sheet_name = output_dic[key][1])
            writer.close()
            
            if graph_flag:
                tk.messagebox.showinfo('Complete', f'File written to:\n{out_file}\n\nGraphs written to:\n{graph_dir}')
            else:
                tk.messagebox.showinfo('Complete', f'File written to:\n\n{out_file}')
            sys.exit()

        elif graph_flag:
            tk.messagebox.showinfo('Complete', f'File written to:\n\n{graph_dir}')
            sys.exit()

        else:
            tk.messagebox.showerror('Unknown Error', 'Big Oopsie\nGraphs are probably written\nSpreadsheets probably aren\'t\n Good Luck!!')
            return

    def save_preset(self):
        while True:
            name = simpledialog.askstring('Preset Name', 'Preset Name:')
            path = self.location + r'\analytics\presets\{}.json'.format(name)
            if os.path.exists(path):
                if tk.messagebox.askyesno('Preset Exists', 'Preset already exists, would you like to overwrite it?'):
                    break
                else:
                    continue
            else:
                break

        out = {"sheets" : {},
               "graphs" : {}}
        
        for idx, key in enumerate(self.spreadsheet_dict):
            out['sheets'][key] = self.spreadsheet_dict[key][1].get()
        
        for idx, key in enumerate(self.graph_dict):
            out['graphs'][key] = self.graph_dict[key][1].get()

        with open(path, 'w') as presetfile:
            json.dump(out, presetfile, indent = 4)

        tk.messagebox.showinfo('Preset Saved', 'Preset saved as {}'.format(name))
        return
    
    def changefile(self):
        newdir = filedialog.askdirectory(title = "Select Output Folder")
        if newdir != '':
            self.out_dir = newdir
            self.file_lbl.config(text = self.out_dir)

            self.proginfo['Analytics Location'] = self.out_dir
            with open(r'gui\config\ProgInfo.json','w') as file:
                json.dump(dict(self.proginfo), file, indent = 4)
  

if __name__ == "__main__":
    print('Analytics Start')
    app = CtrlPanel()
    app.mainloop()