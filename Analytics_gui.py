import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import filedialog
from analytics.funcs import *
from gui.utils import gui_funcs
import json
import pandas as pd
import sys
import os


class CtrlPanel(ThemedTk):
    def __init__(self):
        self.location = os.path.dirname(__file__)
        try:
            with open(self.location + r'\gui\config\ProgInfo.json', 'r') as progfile:
             self.proginfo = json.load(progfile)
        except Exception as err:
            gui_funcs.GenericError(err)
            sys.exit()
        
        self.out_dir = self.proginfo['Analytics Location']

        self.creation_dict = {"basicframe" : ['Standard Stats'],
                         "comboframe" : ['Combinaitons'],
                         "uniqueframe" : ["No. Unique People Sniped"],
                         "firstframe" : ["First Snipe/Sniped Dates"],
                         "dateframe" : ["Total Snipes per Day"],
                         "weekframe" : ["Total Snipes per Week"],
                         "sn_dateframe" : ["Ind. Snipes per Day"],
                         "sd_dateframe": ["Ind. Sniped per Day"],
                         "sn_byweek" : ["Ind. Snipes per Week"],
                         "sd_byweek" : ["Ind. Sniped per Week"],
                         }
        
        ThemedTk.__init__(self)
        self.title("Analytics Control Panel")

        self.sheets_frm = tk.Frame(master = self)
        self.sheets_frm.pack()

        for idx, key in enumerate(self.creation_dict):
            # temp_frm = ttk.Frame(master = self.sheets_frm)
            lbl = ttk.Label(master = self.sheets_frm, text = self.creation_dict[key][0] + " :")
            bool_var = tk.BooleanVar()
            box = ttk.Checkbutton(self.sheets_frm, variable=bool_var)

            self.creation_dict[key].append(bool_var)

            lbl.grid(row = idx, column= 0, sticky = 'w', padx = (0,65), pady = 5)
            box.grid(row = idx, column= 1, sticky = 'e')

        self.submit_frm = ttk.Frame(master = self)
        self.submit_frm.pack()

        self.sub_btn = ttk.Button(master = self.submit_frm, text = "Generate", command = self.btncmd)
        self.sub_btn.pack()

        self.settings_frm = ttk.Frame(master = self, borderwidth = 2, relief = 'sunken')
        self.settings_frm.pack(padx = 4, pady = 4)

        out_lbl = ttk.Label(master = self.settings_frm, text = "Output File Location:")
        out_lbl.grid(row = 0, column = 0, sticky = 'w')

        self.file_lbl = ttk.Label(master = self.settings_frm, text = self.out_dir, width = 40, background = 'darkgray')
        self.file_lbl.grid(row = 1, column = 0, columnspan = 5)
        change_btn = ttk.Button(master = self.settings_frm, text = "Set", command = self.changefile)
        change_btn.grid(row = 1, column = 6)

    def btncmd(self):
        top = tk.Toplevel(self)
        top.title('Working...')

        msgstr = tk.StringVar()
        msgstr.set('Contacting Google...')
        msg = ttk.Label(master = top, textvariable = msgstr)
        msg.pack(padx = 10, pady = 10)

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

        raw, names_dict = import_data(self.creds, self.proginfo['Google Sheet Name'], self.proginfo['User'], self.proginfo['Whitelist Sheet'])
        
        output_dic = {}

        msgstr.set(msgstr.get() + '\nProcessing Data...')
        top.update()
        try:
            if self.creation_dict['basicframe'][1].get():
                output_dic['basicframe'] = [create_basicframe(raw, names_dict), "Basic"]

            if self.creation_dict['comboframe'][1].get():
                output_dic['comboframe'] = [create_comboframe(raw, names_dict), "Combo"]

            if self.creation_dict['uniqueframe'][1].get():
                if "comboframe" in output_dic.keys():
                    output_dic['uniqueframe'] = [create_uniqueframe(output_dic['comboframe'][0], names_dict), "Unique"]
                else:
                    tempcombo = create_comboframe(raw, names_dict)
                    output_dic['uniqueframe'] = [create_uniqueframe(tempcombo, names_dict), "Unique"]
            
            if self.creation_dict['dateframe'][1].get() or self.creation_dict['weekframe'][1].get():
                tempdate, tempweek, truncdateframe, daterange = create_timetotal(raw)
                if self.creation_dict['dateframe'][1].get():
                    output_dic['dateframe'] = [tempdate, "Total by Day"]
                    
                if self.creation_dict['weekframe'][1].get():
                    output_dic['weekframe'] = [tempweek, "Total by Week"]

            if self.creation_dict['sn_dateframe'][1].get() or self.creation_dict['sd_dateframe'][1].get() or self.creation_dict['sn_byweek'][1].get() or self.creation_dict['sd_byweek'][1].get():
                try:
                    truncdateframe
                except NameError:
                    tempdate, tempweek, truncdateframe, daterange = create_timetotal(raw)

                sn_dateframe, sd_dateframe, sn_byweek, sd_byweek = create_timeind(truncdateframe, names_dict, daterange)

                if self.creation_dict['sn_dateframe'][1].get():
                        output_dic['sn_dateframe'] = [sn_dateframe, "Ind. Sn. Day"]

                if self.creation_dict['sd_dateframe'][1].get():
                        output_dic['sd_dateframe'] = [sd_dateframe, "Ind. Sd. Day"]

                if self.creation_dict['sn_byweek'][1].get():
                        output_dic['sn_byweek'] = [sn_byweek, "Ind. Sn. Week"]

                if self.creation_dict['sd_byweek'][1].get():
                        output_dic['sd_byweek'] = [sd_byweek, "Ind. Sd. Week"]
        except Exception as e:
            gui_funcs.GenericError(e)
            return
        

        msgstr.set(msgstr.get() + '\nWriting File...')
        top.update()

        if len(output_dic) > 0:
            out_file = self.out_dir + r'\SnipeStats {}.xlsx'.format(pd.Timestamp.now().strftime('%Y-%m-%d %H-%M'))
            writer = pd.ExcelWriter(out_file, engine='openpyxl')
            for key in output_dic:
                print('hit')
                print(key)
                print(type(output_dic[key][0]))
                print()
                output_dic[key][0].to_excel(writer, sheet_name = output_dic[key][1])
            writer.close()
        
            tk.messagebox.showinfo('Complete', f'File written to:\n\n{out_file}')
            sys.exit()

        else:
             tk.messagebox.showerror('No Options','No options selected\nPlease select some stats')

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