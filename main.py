# region Imports
import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import pandas as pd
import os
import numpy as np
from default_dictionary import default_dict

# xlsxwriter is a soft dependency for export formatting, apparent error with pycharm interp of Pandas ExcelWriter
# endregion


# region Pandas Print Display Options for Testing
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', 15)
pd.set_option('display.width', None)


# endregion


class MainFrame:
    def __init__(self, master):
        master.geometry('950x850')
        master.minsize(950, 850)

        # region Variables
        self.dict_check = False  # Checks if dictionary selection is valid
        self.fp_check = False  # Checks if file path is valid
        self.fp = str()  # Saves import filepath

        # Import data & join dataframes
        self.prime_df = pd.DataFrame()  # Primary dataset
        self.second_df = pd.DataFrame()  # Secondary dataset
        self.tertiary_df = pd.DataFrame()  # Tertiary dataset
        self.support_df = pd.DataFrame()  # Supporting data

        self.joined_df = pd.DataFrame()  # Finalised concatenate joined dataframe for export

        # Internal & imported dictionary dataframes
        self.default_dict = default_dict  # Default dictionary contains all known fields for all test regions
        self.imported_dict = pd.DataFrame()  # Imported dictionary for custom filtering

        # endregion

        # region User Interface

        # region 0.0 - Master Frame
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        self.top_frame = tk.Frame(master)
        self.top_frame.grid(row=0, column=0, sticky='NSEW')
        self.top_frame.columnconfigure(0, weight=1)
        self.top_frame.rowconfigure(3, weight=1)
        self.top_frame.config(bg='#282828', highlightbackground='#FFFFFF')
        # endregion

        # region 1.0 - Select Imports Frame
        self.import_frame = tk.Frame(self.top_frame, padx=5, pady=2)
        self.import_frame.grid(row=0, column=0, sticky='EW')
        self.import_frame.columnconfigure(1, weight=1)

        # 1.1 - Use Default Dictionary Checkbutton
        self.dict_var = tk.IntVar(value=0)
        self.dict_check = tk.Checkbutton(self.import_frame, text='Use Default Mapping', variable=self.dict_var,
                                         command=self.use_default_dict)
        self.dict_check.grid(row=0, column=0, columnspan=2, sticky='W')

        # 1.2 - Use Imported Dictionary Button & Display Label
        self.dict_open_button = tk.Button(self.import_frame, text='Open:', command=self.import_dict, width=10)
        self.dict_open_button.grid(row=1, column=0, sticky='EW')
        self.dict_label = tk.Label(self.import_frame, text='Select Mapping File (.txt or .xlsx)', anchor='w')
        self.dict_label.grid(row=1, column=1, columnspan=2, sticky='EW')

        # 1.3 - Select Filepath Button & Label Display
        self.import_open_button = tk.Button(self.import_frame, text='Open:', command=self.import_path, width=10)
        self.import_open_button.grid(row=2, column=0, sticky='EW')
        self.import_label = tk.Label(self.import_frame, text='Select Data Import Directory', anchor='w')
        self.import_label.grid(row=2, column=1, columnspan=2, sticky='EW')

        # 1.5 - Import Data Button
        self.import_load_button = tk.Button(self.import_frame, text='Import Data', command=self.prelim_load,
                                            width=10, state='disabled')
        self.import_load_button.grid(row=1, rowspan=4, column=2, sticky='NSE')
        # endregion

        # region 2.0 - File Contents Frame
        self.dir_frame = tk.Frame(self.top_frame, padx=5, pady=5)
        self.dir_frame.grid(row=2, column=0, sticky='EW')

        # 2.1 - Y Axis Headers
        for y, header in enumerate(['Files', 'Files\nList', 'Lines', 'Cols.']):
            tk.Label(self.dir_frame, text=header).grid(row=y + 1, column=0, sticky='W')

        # 2.2 - Import GUI Generation
        self.widget_list = []
        for x, header in enumerate(['Primary', 'Secondary', 'Tertiary', 'Supporting', 'Other']):
            # 2.3 - X Axis Headers
            tk.Label(self.dir_frame, text=header).grid(row=0, column=x + 1)
            self.dir_frame.columnconfigure(x + 1, weight=1)

            # 2.4 - File Number
            self.file_no = tk.Entry(self.dir_frame, justify='center')
            self.file_no.grid(row=1, column=x + 1, sticky='EW', pady=2)

            # 2.5 - File List
            self.file_list = tk.Listbox(self.dir_frame, height=7)
            self.file_list.grid(row=2, column=x + 1, sticky='EW')

            # 2.6 - File Lines
            self.file_lines = tk.Entry(self.dir_frame, justify='center')
            self.file_lines.grid(row=3, column=x + 1, sticky='EW', pady=2)

            # 2.7 - File Columns
            self.file_cols = tk.Entry(self.dir_frame, justify='center')
            self.file_cols.grid(row=4, column=x + 1, sticky='EW', pady=2)

            # 2.8 - Populate widget list & Disable Other
            self.widget_list.append([self.file_no, self.file_list, self.file_lines, self.file_cols])

        # endregion

        # region 3.0 - Join, Concat, Drop Data - Master Frame
        self.data_frame = tk.Frame(self.top_frame)
        self.data_frame.grid(row=3, column=0, sticky='NSEW')
        self.data_frame.columnconfigure(1, weight=1)
        self.data_frame.rowconfigure(0, weight=1)
        # endregion

        # region 4.0 - Join, Concat, Drop Data - Options
        self.opt_frame = tk.Frame(self.data_frame, padx=5, pady=5)
        self.opt_frame.grid(row=0, column=0, sticky='NSEW')

        # 4.1 - Options Header
        self.data_label = tk.Label(self.opt_frame, text='Filter Options:')
        self.data_label.grid(column=0, sticky='W')

        # 4.2 - Total Generation Options
        self.data_total_label = tk.Label(self.opt_frame, text='Generate Total Fields:')
        self.data_total_label.grid(column=0, sticky='W', pady=5)

        # Primary Data consolidation
        self.prime_var = tk.IntVar(value=1)
        self.prime_check = tk.Checkbutton(self.opt_frame, text='Primary Sub-Totals',
                                          variable=self.prime_var, command=self.drop_prime)
        self.prime_check.grid(column=0, sticky='W')
        self.prime_drop_var = tk.IntVar(value=0)
        self.prime_drop_check = tk.Checkbutton(self.opt_frame, text='Drop Sub-Totals', variable=self.prime_drop_var)
        self.prime_drop_check.grid(column=0)

        # Tertiary Data consolidation
        self.tert_var = tk.IntVar(value=1)
        self.tert_check = tk.Checkbutton(self.opt_frame, text='Tertiary Sub-Totals',
                                         variable=self.tert_var, command=self.drop_tert)
        self.tert_check.grid(column=0, sticky='W')
        self.tert_drop_var = tk.IntVar(value=0)
        self.tert_drop_check = tk.Checkbutton(self.opt_frame, text='Drop Sub-Totals', variable=self.tert_drop_var)
        self.tert_drop_check.grid(column=0)

        # 4.3 - Consolidate Fields
        self.data_consol_label = tk.Label(self.opt_frame, text='Consolidate Fields:')
        self.data_consol_label.grid(column=0, sticky='W', pady=5)

        # Generates checkbutton list
        self.checkbutton_list = []
        for field in ['Area', 'Timepoint (M)', 'Timepoint (D)', 'Test No.', 'M/F', 'Prelim. Score', 'Post. Score']:
            self.var = tk.IntVar(value=1)
            self.check = tk.Checkbutton(self.opt_frame, text=field, variable=self.var)
            self.check.grid(column=0, sticky='W')
            self.checkbutton_list.append([field, self.var])

        self.opt_frame_2 = tk.Frame(self.opt_frame, pady=10)
        self.opt_frame_2.grid(column=0, sticky='EW')

        self.drop_var = tk.IntVar(value=0)
        self.drop_originals = tk.Checkbutton(self.opt_frame_2, text='Drop Original Fields', variable=self.drop_var)
        self.drop_originals.grid(column=0, sticky='W')

        self.data_consol_button = tk.Button(self.opt_frame, text='Join Data:', height=2, command=self.join_data,
                                            state='disabled')
        self.data_consol_button.grid(column=0, sticky='EW')
        # endregion

        # region 5.0 - Join, Concat, Drop Data - Output Listbox
        self.output_subframe = tk.Frame(self.data_frame, padx=5, pady=5)
        self.output_subframe.grid(row=0, column=1, sticky='NSEW')
        self.output_subframe.columnconfigure(0, weight=1)
        self.output_subframe.rowconfigure(0, weight=1)
        self.data_listbox = tk.Listbox(self.output_subframe)
        self.data_listbox.grid(column=0, sticky='NSEW')
        # endregion

        # region 6.0 - Export Frame
        self.export_frame = tk.Frame(self.top_frame, padx=5, pady=5)
        self.export_frame.grid(row=5, column=0, sticky='EW')
        self.export_frame.columnconfigure(1, weight=1)

        self.export_var = tk.IntVar(value=0)
        self.export_check = tk.Checkbutton(self.export_frame, text='Use Import Directory', variable=self.export_var,
                                           command=self.export_auto_fp, state='disabled')
        self.export_check.grid(row=0, column=0, columnspan=2, sticky='W')

        self.export_fp_button = tk.Button(self.export_frame, text='Open:', command=self.export_path,
                                          height=2, width=10, state='disabled')
        self.export_fp_button.grid(row=1, column=0, sticky='EW')

        self.export_label = tk.Label(self.export_frame, text='Select Export Directory', anchor='w')
        self.export_label.grid(row=1, column=1, sticky='EW')

        self.export_load_button = tk.Button(self.export_frame, text='Export Data', command=self.export_data,
                                            height=2, width=10, state='disabled')
        self.export_load_button.grid(row=1, column=2, sticky='E')
        # endregion

        self.colour(self.top_frame)
        # endregion

    # region Colour Scheme Function
    def colour(self, parent):
        for child in parent.winfo_children():
            widget_type = child.winfo_class()

            if widget_type == 'Frame':
                child.config(bg='#404040')
            if widget_type == 'Label':
                if child.winfo_parent() in ['.!frame.!frame', '.!frame.!frame4']:
                    child.config(bg='#282828', fg='#FFFFFF')
                else:
                    child.config(bg='#404040', fg='#FFFFFF')
            if widget_type == 'Button':
                child.config(bg='#282828', fg='#FFFFFF')
            if widget_type == 'Canvas':
                child.config(bg='#282828')
            if widget_type == 'Checkbutton':
                child.config(bg='#404040', selectcolor='#404040', fg='#FFFFFF', bd=3,
                             activebackground='#404040', activeforeground='#FFFFFF')
            if widget_type == 'Entry':
                child.config(bg='#282828', fg='#FFFFFF', disabledbackground='#282828')
            if widget_type == 'Listbox':
                if child.winfo_parent() == '.!frame.!frame2':
                    if str(child) == '.!frame.!frame2.!listbox5':
                        child.config(bg='#282828', fg='red', highlightthickness=0, font=('Courier', 8))
                    else:
                        child.config(bg='#282828', fg='green', highlightthickness=0, font=('Courier', 8))

                else:
                    child.config(bg='#282828', fg='#FFFFFF', highlightthickness=0, font=('Courier', 8))
            else:
                self.colour(child)

    # endregion

    # region File Path & Dictionary Functions
    def use_default_dict(self):
        if self.dict_var.get() == 1:
            # Disable dictionary import widgets, defer to default dictionary
            self.dict_open_button.config(state='disabled')
            self.dict_label.config(text='Default internal mappings will be used', fg='green')
            self.dict_check = True
            if self.fp_check and self.dict_check:
                self.import_load_button.config(state='normal', fg='yellow')
            else:
                self.import_load_button.config(state='disabled')
        else:
            # Enable dictionary import widgets, defer to default dictionary
            self.dict_open_button.config(state='normal')
            self.dict_label.config(text='Select Mapping File (.txt or .xlsx)', fg='red')
            self.dict_check = False
            self.import_load_button.config(state='disabled')

    def import_dict(self):
        dict_file = askopenfilename()

        try:
            if dict_file.endswith('.xlsx'):
                imported_dict = pd.read_excel(dict_file, sheet_name=0, index_col='Column')
                self.imported_dict = imported_dict['Mapping'].to_dict()
                self.dict_label.config(text='.xlsx file has been converted to dictionary format', fg='green')
                self.dict_check = True
            elif dict_file.endswith('.txt'):
                with open(dict_file, 'r') as file:
                    imported_dict = eval(file.read())
                self.imported_dict = imported_dict
                self.dict_label.config(text='.txt file has been converted to dictionary format', fg='green')
                self.dict_check = True
            else:
                self.dict_label.config(text='File selected does not have an .xlsx or .txt extension', fg='red')
                self.dict_check = False
                return

            if self.fp_check and self.dict_check:
                self.import_load_button.config(state='normal', fg='yellow')
            else:
                self.import_load_button.config(state='disabled')
        except Exception as error:
            self.dict_label.config(text='ERROR:' + str(error), fg='red')
            self.dict_check = False
            return

    def import_path(self):
        fp = filedialog.askdirectory(initialdir='/', title='Select a directory')

        # Sets GUI label text to selected filepath, exits function if not selected
        try:
            if len(fp) != 0:
                self.import_label.config(text=str(fp), fg='green')
                self.fp_check = True
                os.chdir(fp)
                self.fp = fp
                if self.fp_check and self.dict_check:
                    self.import_load_button.config(state='normal', fg='yellow')
                else:
                    self.import_load_button.config(state='disabled')
            else:
                self.import_label.config(text='ERROR: No Filepath Selected', fg='red')
                self.import_load_button.config(state='disabled')
                self.fp_check = False
                return
        except Exception as fp_error:
            self.import_label.config(text='ERROR: ' + str(fp_error))
            self.import_label.config(fg='red')
            self.import_load_button.config(state='disabled')
            self.fp_check = False
            return

    # endregion

    # region Data Load & Join Functions
    def prelim_load(self):
        # 1) Clear Dataframes and List-boxes, disable GUI
        for df in [self.prime_df, self.second_df, self.tertiary_df, self.support_df]:
            df.drop(df.index, inplace=True)
        for lb in self.widget_list:
            lb[1].delete(0, tk.END)
        self.data_consol_button.config(state='disabled')
        self.export_fp_button.config(state='disabled')
        self.export_load_button.config(state='disabled')

        # 2) Assign dictionary to local variable
        dictionary = self.imported_dict if self.dict_var.get() == 0 else self.default_dict

        # 3) Set file import directory & generate import file list
        os.chdir(self.fp)
        import_files = [file for file in os.listdir() if file.endswith('.xlsx') and not file.startswith('~$')
                        and not file == 'Site Data Joined.xlsx'  # Prevents reloading joined data during testing
                        ]

        # 4) Import data, assign to dataframes
        for file in import_files:
            try:
                # 4.1) Import only columns present in dictionary
                data = pd.read_excel(file, sheet_name=0, usecols=lambda x: x in set(dictionary))

                # 4.2) If data has less than 2 valid columns, skip lines to account for null header lines
                row_skip = 0
                while len(data.columns) < 2:
                    if row_skip < 5:  # Allow loop breakout to prevent infinite reloads
                        row_skip += 1
                        data = pd.read_excel(file, sheet_name=0, usecols=lambda x: x in set(dictionary),
                                             skiprows=row_skip)
                    else:
                        break

                # 4.3) Standardise Data Headers
                data.rename(columns=dictionary, inplace=True)

                # 4.4) Check for multiple ID fields, delete all but first instance as this is always the priority ID
                data = data.loc[:, ~data.columns.duplicated()].copy()

                # 4.5) Drop blanks from ID field, drop entirely blank rows of data
                if 'Reference Code' in data:
                    data.dropna(subset=['Reference Code'], how='any', inplace=True)
                    data.dropna(axis=1, how='all', inplace=True)
                else:
                    self.widget_list[4][1].insert(tk.END, str(file))
                    continue

                # 4.6) Standardise values to int if numeric
                data['Reference Code'] = pd.to_numeric(data['Reference Code'], errors='ignore', downcast='integer')

                # 4.8) Concatenate data based off presence of relevant fields unique to each data type
                prime_check = [col for col in data.columns if 'SUB_' in col]
                tert_check = [col for col in data.columns if 'M-L' in col]
                tert_check_2 = [col for col in data.columns if 'O-L' in col]

                if any(col in data.columns for col in ['Resp. Score', 'Art. Score']) or len(prime_check) != 0:
                    self.prime_df = pd.concat([self.prime_df, data], axis=0)
                    self.widget_list[0][1].insert(tk.END, str(file))
                elif any(col in data.columns for col in ['Sec. Resp. Score', 'Sec. Art. Score']):
                    self.second_df = pd.concat([self.second_df, data], axis=0)
                    self.widget_list[1][1].insert(tk.END, str(file))
                elif any(col in data.columns for col in ['Voc. Group']) or len(tert_check) != 0 \
                        or len(tert_check_2) != 0:
                    self.tertiary_df = pd.concat([self.tertiary_df, data])
                    self.widget_list[2][1].insert(tk.END, str(file))
                elif any(col in data.columns for col in ['Post. Score']):
                    self.support_df = pd.concat([self.support_df, data])
                    self.widget_list[3][1].insert(tk.END, str(file))
                else:
                    self.widget_list[4][1].insert(tk.END, str(file))
            except Exception as error:
                self.widget_list[4][1].insert(tk.END, str(file) + ' - ' + str(error))

        # 5) Update Import GUI
        # File No. & Name count GUI
        for widget in self.widget_list:
            # Add File no. for each data type
            widget[0].delete(0, tk.END)
            widget[0].insert(0, str(widget[1].size()))

            # Updates File No. widget colours, exception for 5th column with "other" files
            if widget[0] != self.widget_list[-1][0]:
                widget[0].config(fg='red') if widget[1].size() == 0 else widget[0].config(fg='green')
            else:
                widget[0].config(fg='red') if widget[1].size() != 0 else widget[0].config(fg='green')
        # Line & Column count GUI
        for lines, widget in zip([self.prime_df, self.second_df, self.tertiary_df, self.support_df], self.widget_list):
            # Add Line counts for each data type
            widget[2].delete(0, tk.END)
            widget[2].insert(0, str(len(lines)))
            widget[2].config(fg='red') if len(lines) == 0 else widget[2].config(fg='green')

            # Add column count for each data type
            widget[3].delete(0, tk.END)
            cols = len(lines.columns)
            widget[3].insert(0, str(cols))
            widget[3].config(fg='red') if cols == 0 else widget[3].config(fg='yellow') if cols > 20 \
                else widget[3].config(fg='green')

        # 5) Enable Import GUI
        self.data_listbox.delete(0, tk.END)
        self.data_listbox.config(fg='#FFFFFF')
        non_prime_len = len(self.second_df) + len(self.tertiary_df) + len(self.support_df)
        if len(self.prime_df) != 0 and non_prime_len != 0:
            # self.tp_entry.config(state='disabled')
            self.import_load_button.config(fg='green')
            self.data_consol_button.config(state='normal', fg='yellow')
            self.data_listbox.insert(tk.END, 'Primary & Additional Data imported.')
            self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='green')
            self.data_listbox.insert(tk.END, '')
            prime_len = len(self.prime_df.columns)
            tert_len = len(self.tertiary_df.columns)
            if prime_len > 100 or tert_len > 100:
                self.data_listbox.insert(tk.END, 'WARNING:')
                self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                if prime_len > 100:
                    self.data_listbox.insert(tk.END, 'Prime data contains ' + str(prime_len) + ' cols.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                    self.data_listbox.insert(tk.END, 'Large volume of Prime data cols. suggests some prime data files '
                                                     'may only contain sub-totals.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                    self.data_listbox.insert(tk.END, 'Advise setting "Generate Totals" & "Drop Sub-Totals" options to '
                                                     'auto gen. a custom total field if data has been checked for '
                                                     'consistency')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                    self.data_listbox.insert(tk.END, '')
                if tert_len > 20:
                    self.data_listbox.insert(tk.END, 'Tertiary data contains ' + str(tert_len) + ' cols.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                    self.data_listbox.insert(tk.END, 'Large volume of Tertiary cols. suggests a mixture of formats, '
                                                     'e.g. sub-totals, totals, comments.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                    self.data_listbox.insert(tk.END, 'Advise setting "Generate Totals" & "Drop Sub-Totals" options to '
                                                     'consolidate and standardise amounts if data has been checked for '
                                                     'consistency')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                    self.data_listbox.insert(tk.END, '')
        else:
            self.data_consol_button.config(state='disabled')
            if len(self.prime_df) != 0 and non_prime_len == 0:
                self.data_listbox.insert(tk.END, 'ERROR: No Additional Data to join from.')
            elif len(self.prime_df) == 0 and non_prime_len != 0:
                self.data_listbox.insert(tk.END, 'ERROR: No Primary Data to join to.')
            else:
                self.data_listbox.insert(tk.END, 'ERROR: No Importable Data found')
            self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='red')

    def drop_prime(self):
        self.prime_drop_var.set(0)
        self.prime_drop_check.config(state='normal') if self.prime_var.get() == 1 \
            else self.prime_drop_check.config(state='disabled')

    def drop_tert(self):
        self.tert_drop_var.set(0)
        self.tert_drop_check.config(state='normal') if self.tert_var.get() == 1 \
            else self.tert_drop_check.config(state='disabled')

    def join_data(self):
        self.export_fp_button.config(state='disabled')
        self.export_check.config(state='disabled')
        self.export_load_button.config(state='disabled')

        join_gui = []

        # 1) Copy prime data to local dataframe, generate Primary data total field
        df1_final = self.prime_df.copy()
        prime_sum_cols = []
        if self.prime_var.get() == 1:
            if not self.prime_df.empty:
                # 1.1) Get list of relevant columns
                prime_sum_cols = [col for col in self.prime_df if col.startswith('SUB_')]
                if len(prime_sum_cols) != 0:
                    # 1.2) Standardise subtotal fields to numeric values
                    df1_final[prime_sum_cols] = df1_final[prime_sum_cols].replace(' ', np.nan)
                    df1_final[prime_sum_cols] = df1_final[prime_sum_cols].fillna(0, axis=0)

                    # 1.3) Generate total fields and drop subtotal fields
                    df1_final['Resp. Score AT '] = df1_final[prime_sum_cols].astype(int).sum(axis=1)
                    df1_final['Art. Score AT '] = (df1_final[prime_sum_cols].astype(int) == 2).sum(axis=1) * 2
                    if self.prime_drop_var.get() == 1:
                        df1_final.drop(columns=prime_sum_cols, axis=1, inplace=True)
            else:
                prime_sum_cols = []

        # 2) Copy tertiary data to local dataframe, generate Primary data total field
        df3_final = self.tertiary_df.copy()
        tert_cols = []
        if self.tert_var.get() == 1:
            if not self.tertiary_df.empty:
                # 2.0) Consolidate all lang Comments into single column including their original source, break by line
                com_cols = sorted([col for col in self.tertiary_df if 'Note' in col])
                if len(com_cols) != 0:
                    tert_cols.append('Merged M/O Notes' + ' (' + str(len(com_cols)) + ' cols. found)')
                    df3_final['Merged M/O Notes'] = np.NaN
                    for col in com_cols:
                        df3_final['Merged M/O Notes'] = np.where(df3_final['Merged M/O Notes'].isnull().all(),
                                                                 str(col) + ': ' + df3_final[col].astype(str),
                                                                 df3_final['Merged M/O Notes'].astype(str) + '\n' +
                                                                 str(col) + ': ' + df3_final[col].astype(str))
                        if self.tert_drop_var.get() == 1:
                            df3_final.drop(col, axis=1, inplace=True)

                # 2.1) Consolidate Main Language Subtotal columns
                ml_sub_totals = sorted([col for col in self.tertiary_df if 'M-L' in col
                                        and any(subtotal in col for subtotal in [' Test '])])
                if len(ml_sub_totals) != 0:
                    df3_final['M-L Test AT '] = 0
                    tert_cols.append('M-L Test AT ' + ' (' + str(len(ml_sub_totals)) + ' cols. found)')
                    for col in ml_sub_totals:
                        # Remove common string characters
                        df3_final[str(col) + '_NUM'] = \
                            df3_final[col].astype(str).str.replace(r'[a-zA-Z%()%<>]', '', regex=True)

                        # Check if characters can be converted to numeric
                        df3_final[str(col) + '_NUM'] = pd.to_numeric(df3_final[str(col) + '_NUM'], errors='coerce')

                        # Convert decimal percentages to full values
                        df3_final[str(col) + '_NUM'] = \
                            df3_final[str(col) + '_NUM'].apply(lambda perc: perc * 100 if perc <= 1 else perc)

                        # Replace NaN's with 0
                        df3_final[str(col) + '_NUM'] = df3_final[str(col) + '_NUM'].fillna(0, axis=0)

                        # Add values to autogenerated total field
                        df3_final['M-L Test AT '] = (df3_final['M-L Test AT '] + df3_final[str(col) + '_NUM'])

                        # Drop subtotal calculation field
                        df3_final.drop([str(col) + '_NUM'], axis=1, inplace=True)

                        if self.tert_drop_var.get() == 1:
                            df3_final.drop(col, axis=1, inplace=True)

                # 2.2) Consolidate Other Language SubTotal columns
                ol_sub_totals = sorted([col for col in self.tertiary_df if 'O-L' in col
                                        and any(subtotal in col for subtotal in [' Test '])])
                if len(ol_sub_totals) != 0:
                    df3_final['O-L Test AT (subs)'] = 0
                    tert_cols.append('O-L Test AT (subs)' + ' (' + str(len(ol_sub_totals)) + ' cols. found)')
                    for col in ol_sub_totals:
                        # Remove common string characters
                        df3_final[str(col) + '_NUM'] = \
                            df3_final[col].astype(str).str.replace(r'[a-zA-Z%()%<>]', '', regex=True)

                        # Check if characters can be converted to numeric
                        df3_final[str(col) + '_NUM'] = pd.to_numeric(df3_final[str(col) + '_NUM'], errors='coerce')

                        # Convert decimal percentages to full values
                        df3_final[str(col) + '_NUM'] = \
                            df3_final[str(col) + '_NUM'].apply(lambda perc: perc * 100 if perc <= 1 else perc)

                        # Replace NaN's with 0
                        df3_final[str(col) + '_NUM'] = df3_final[str(col) + '_NUM'].fillna(0, axis=0)

                        # Add values to autogenerated total field
                        df3_final['O-L Test AT (subs)'] = (
                                df3_final['O-L Test AT (subs)'] + df3_final[str(col) + '_NUM'])

                        # Drop subtotal calculation field
                        df3_final.drop([str(col) + '_NUM'], axis=1, inplace=True)

                        if self.tert_drop_var.get() == 1:
                            df3_final.drop(col, axis=1, inplace=True)

                    tert_cols.append('Voc. Group AT (subs)')
                    df3_final['Voc. Group AT (subs)'] = df3_final['O-L Test AT (subs)'].apply(
                        lambda perc: 1 if perc >= 20 else 0)

                # 2.3) Consolidate Other Language Total columns
                ol_totals = sorted([col for col in self.tertiary_df if 'O-L' in col
                                    and any(subtotal in col for subtotal in ['Total'])])
                if len(ol_totals) != 0:
                    df3_final['O-L Test AT (tot.)'] = 0
                    tert_cols.append('O-L Test AT (tot.)' + ' (' + str(len(ol_sub_totals)) + ' cols. found)')
                    for col in ol_totals:
                        # Remove common string characters
                        df3_final[str(col) + '_NUM'] = \
                            df3_final[col].astype(str).str.replace(r'[a-zA-Z%()%<>]', '', regex=True)

                        # Check if characters can be converted to numeric
                        df3_final[str(col) + '_NUM'] = pd.to_numeric(df3_final[str(col) + '_NUM'], errors='coerce')

                        # Convert decimal percentages to full values
                        df3_final[str(col) + '_NUM'] = \
                            df3_final[str(col) + '_NUM'].apply(lambda perc: perc * 100 if perc <= 1 else perc)

                        # Replace NaN's with 0
                        df3_final[str(col) + '_NUM'] = df3_final[str(col) + '_NUM'].fillna(0, axis=0)

                        # Add values to autogenerated total field
                        df3_final['O-L Test AT (tot.)'] = (
                                df3_final['O-L Test AT (tot.)'] + df3_final[str(col) + '_NUM'])

                        # Drop subtotal calculation field
                        df3_final.drop([str(col) + '_NUM'], axis=1, inplace=True)

                        if self.tert_drop_var.get() == 1:
                            df3_final.drop(col, axis=1, inplace=True)

                    tert_cols.append('Voc. Group AT (tot.)')
                    df3_final['Voc. Group AT (tot.)'] = df3_final['O-L Test AT (tot.)'].apply(
                        lambda perc: 1 if perc >= 20 else 0)

        # 3) Join Datasets
        joined_data = df1_final
        joined_data['Reference Code'] = joined_data['Reference Code'].astype(str)
        for join, suffix in zip([self.second_df, df3_final, self.support_df], [' - SCND', ' - TERT', ' - SUPP']):
            join_initial_len = len(joined_data)
            if not join.empty:
                join['Reference Code'] = join['Reference Code'].astype(str)
                joined_data = joined_data.merge(join, how='outer', on='Reference Code', suffixes=('', suffix))
            join_gui.append([join_initial_len, len(join), len(joined_data)])

        # 4) Set ID as index and standardise
        def col_order(word):
            if word[:4] == 'SUB_':
                return 100
            else:
                try:
                    col_order_dict = {'Reference Code': 1,
                                      'Area': 2,
                                      'Area - SCND': 3,
                                      'Area - SUPP': 5,
                                      'Area - TERT': 4,
                                      'Area - MERGE': 6,
                                      'Timepoint (D)': 7,
                                      'Timepoint (D) - SCND': 8,
                                      'Timepoint (D) - TERT': 9,
                                      'Timepoint (M)': 11,
                                      'Timepoint (M) - SCND': 12,
                                      'Timepoint (M) - SUPP': 14,
                                      'Timepoint (M) - TERT': 13,
                                      'Timepoint (M) - MERGE': 15,
                                      'Timepoint (D) - MERGE': 10,
                                      'Test No.': 16,
                                      'Test No. - SUPP': 17,
                                      'Test No. - MERGE': 18,
                                      'M/F': 19,
                                      'M/F - SCND': 20,
                                      'M/F - SUPP': 22,
                                      'M/F - TERT': 21,
                                      'M/F - MERGE': 23,
                                      'F. Qual.': 24,
                                      'M. Qual.': 25,
                                      'Prelim. Score': 26,
                                      'Prelim. Score - SCND': 27,
                                      'Prelim. Score - SUPP': 29,
                                      'Prelim. Score - TERT': 28,
                                      'Prelim. Score - MERGE': 30,
                                      'Resp. Score': 31,
                                      'Resp. Score AT ': 32,
                                      'Art. Score': 33,
                                      'Art. Score AT ': 34,
                                      'Sec. Art. Score': 39,
                                      'Sec. Art. Score (Scaled)': 41,
                                      'Sec. Art. Score (Scaled) - SCND': 42,
                                      'Sec. Art. Score - SCND': 40,
                                      'Sec. Resp. Score': 35,
                                      'Sec. Resp. Score (Scaled)': 37,
                                      'Sec. Resp. Score (Scaled) - SCND': 38,
                                      'Sec. Resp. Score - SCND': 36,
                                      'Post. Score': 43,
                                      'Post. Score - SUPP': 44,
                                      'Post. Score - MERGE': 45,
                                      'M-L Note': 46,
                                      'M-L Test 1': 51,
                                      'M-L Test 2': 52,
                                      'M-L Test AT ': 54,
                                      'M-L Total': 53,
                                      'Merged M/O Notes': 50,
                                      'O-L Test AT (subs)': 61,
                                      'O-L Test AT (tot.)': 65,
                                      'O-L Total': 64,
                                      'O-L1  Note': 47,
                                      'O-L1 Test 1': 55,
                                      'O-L1 Test 2': 56,
                                      'O-L1 Total': 62,
                                      'O-L2  Note': 48,
                                      'O-L2 Test 1': 57,
                                      'O-L2 Test 2': 58,
                                      'O-L2 Total': 63,
                                      'O-L3  Note': 49,
                                      'O-L3  Test 1': 59,
                                      'O-L3  Test 2': 60,
                                      'Voc. Group': 61,
                                      'Voc. Group AT (subs)': 62,
                                      'Voc. Group AT (tot.)': 63}

                    return col_order_dict[word]

                except KeyError:
                    return 100

        joined_data.set_index('Reference Code', inplace=True)
        joined_data = joined_data.reindex(sorted(joined_data.columns), axis=1)

        # 5) Consolidate selected fields into single column, prioritizing from prime > supporting
        convert_gui = []
        # Get list of fields to merge from GUI selections
        merge_selections = [selections[0] for selections in self.checkbutton_list if selections[1].get() == 1]

        for col_name in merge_selections:  # Loop through each selection
            cols_found = [substring for substring in joined_data if col_name in substring]  # Get cols with substring
            convert_gui.append([col_name, len(cols_found)])  # Add field & no. of cols to GUI for output to listbox

            if len(cols_found) > 1:
                joined_data[str(col_name) + str(' - MERGE')] = joined_data[col_name]  # Create custom _MERGE column
                for col in cols_found:  # Iterate from Primary > Supporting priority order filling in blank data
                    joined_data[str(col_name) + str(' - MERGE')] = \
                        joined_data[str(col_name) + str(' - MERGE')].fillna(joined_data[col])  # Fill where null

                if self.drop_var.get() == 1:  # Drop original fields if selected
                    joined_data.drop(cols_found, axis=1, inplace=True)

        joined_data = joined_data.reindex(sorted(joined_data.columns, key=col_order), axis=1)
        joined_data.dropna(how='all', axis=1, inplace=True)

        self.joined_df = joined_data
        self.data_consol_button.config(fg='green')
        self.export_fp_button.config(state='normal')
        self.export_check.config(state='normal')

        # region 6) Update Listbox GUI & Enable Export GUI
        self.data_listbox.delete(0, tk.END)
        self.data_listbox.config(fg='#FFFFFF')

        self.data_listbox.insert(tk.END, '=====================================================')
        self.data_listbox.insert(tk.END, '============= 1) Total Field Generation =============')
        self.data_listbox.insert(tk.END, '=====================================================')
        self.data_listbox.insert(tk.END, '')

        if self.prime_var.get() == 1:
            self.data_listbox.insert(tk.END, '------------- Primary Totals ------------')
            if len(prime_sum_cols) != 0:
                self.data_listbox.insert(tk.END, '')
                self.data_listbox.insert(tk.END, str(len(prime_sum_cols)) + ' Primary Data subtotals found, '
                                                                            'following total fields generated:')
                self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='green')
                self.data_listbox.insert(tk.END, '')
                self.data_listbox.insert(tk.END, '- Resp. Score AT')
                self.data_listbox.insert(tk.END, '- Art. Score AT')
                self.data_listbox.insert(tk.END, '')

                # Standardise subtotal fields so they can be summed
                if self.prime_drop_var.get() == 1:
                    self.data_listbox.insert(tk.END, 'Primary Dataset Subtotal fields have been dropped.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='green')
                else:
                    self.data_listbox.insert(tk.END, 'Primary Dataset Subtotal fields have been retained.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
            else:
                self.data_listbox.insert(tk.END, 'No standard subtotal fields found in Primary Datasets')
                self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
        self.data_listbox.insert(tk.END, '')

        if self.tert_var.get() == 1:
            self.data_listbox.insert(tk.END, '------------ Tertiary Totals ------------')
            if len(tert_cols) != 0:
                self.data_listbox.insert(tk.END, '')
                self.data_listbox.insert(tk.END, 'Multiple Tertiary Data subtotals found, '
                                                 'following total fields generated:')
                self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='green')
                self.data_listbox.insert(tk.END, '')
                for col in tert_cols:
                    self.data_listbox.insert(tk.END, '- ' + str(col))
                self.data_listbox.insert(tk.END, '')

                # Standardise subtotal fields so they can be summed
                if self.tert_drop_var.get() == 1:
                    self.data_listbox.insert(tk.END, 'Tertiary Dataset Subtotal fields have been dropped.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='green')
                else:
                    self.data_listbox.insert(tk.END, 'Tertiary Dataset Subtotal fields have been retained.')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
            else:
                self.data_listbox.insert(tk.END, 'No standard subtotal fields found in Tertiary datasets')
                self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
        self.data_listbox.insert(tk.END, '')

        if self.prime_var.get() != 1 and self.tert_var.get() != 1:
            self.data_listbox.insert(tk.END, 'No subtotal consolidation options selected.')
            self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
            self.data_listbox.insert(tk.END, '')

        self.data_listbox.insert(tk.END, '=====================================================')
        self.data_listbox.insert(tk.END, '================== 2) Dataset Joins =================')
        self.data_listbox.insert(tk.END, '=====================================================')
        self.data_listbox.insert(tk.END, '')

        for join_name, join_len in zip(['Secondary', 'Tertiary', 'Supporting'], join_gui):
            if join_len[1] != 0:
                self.data_listbox.insert(tk.END, '------------- ' + join_name + ' Join ------------')
                self.data_listbox.insert(tk.END, '')
                self.data_listbox.insert(tk.END, 'Initial Data Len:       ' + str(join_len[0]) + ' Lines')
                spacer = ''
                for x in range(23 - len(join_name)):
                    spacer += ' '
                self.data_listbox.insert(tk.END, join_name + ':' + spacer + str(join_len[1]) + ' Lines')
                self.data_listbox.insert(tk.END, 'New Data Len:           ' + str(join_len[2]) + ' Lines')
                self.data_listbox.insert(tk.END, '')

                if join_len[0] != join_len[2]:
                    null_codes = join_len[2] - join_len[0]
                    self.data_listbox.insert(tk.END, 'NOTE: ' + join_name + ' data contains ' + str(null_codes) +
                                             ' join IDs with no corresponding Primary data')
                    self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
                    self.data_listbox.insert(tk.END, '')

        self.data_listbox.insert(tk.END, '=====================================================')
        self.data_listbox.insert(tk.END, '============== 3) Field Consolidation ===============')
        self.data_listbox.insert(tk.END, '=====================================================')
        self.data_listbox.insert(tk.END, '')

        if len(convert_gui) != 0:
            for field, length in convert_gui:
                spacer = ''
                for x in range(22 - len(field)):
                    spacer += ' '
                if length != 0:
                    self.data_listbox.insert(tk.END, field + ': ' + spacer + str(length) + ' fields consolidated')
                else:
                    self.data_listbox.insert(tk.END, field + ': ' + spacer + 'No subtotal fields found to consolidate')
        else:
            self.data_listbox.insert(tk.END, 'No field consolidation selections were made')

            self.data_consol_button.config(fg='green')

        self.data_listbox.insert(tk.END, '')
        self.data_listbox.insert(tk.END, 'Joined data Lines:      ' + str(len(joined_data)))
        self.data_listbox.insert(tk.END, 'Joined data Columns:    ' + str(len(joined_data.columns)))
        self.data_listbox.insert(tk.END, '')

        dupe_list = set(ref for ref in list(joined_data.index) if list(joined_data.index).count(ref) > 1)
        self.data_listbox.insert(tk.END, 'Duplicated Reference Codes:    ' + str(len(dupe_list)))
        if len(dupe_list) == 0:
            self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='green')
        else:
            self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')
        self.data_listbox.insert(tk.END, '')
        for code in dupe_list:
            self.data_listbox.insert(tk.END, code)
            self.data_listbox.itemconfig(self.data_listbox.size() - 1, fg='yellow')

        for child in self.export_frame.winfo_children():
            widget_type = child.winfo_class()
            if widget_type == 'Button':
                child: tk.Button
                if child == '.!frame.!frame4.!button':
                    child.config(state='normal')
                if child == '.!frame.!frame4.!button2':
                    child.config(state='normal', fg='yellow')
            if widget_type == 'Checkbutton':
                child: tk.Checkbutton
                child.config(state='normal')
        # endregion

    # endregion

    # region Export Functions
    def export_auto_fp(self):
        if self.export_var.get() == 1:
            self.export_fp_button.config(state='disabled')
            if len(self.joined_df) != 0:
                self.export_load_button.config(state='normal')
                fp = self.import_label['text']
                self.export_label.config(text=fp, fg='green')
            else:
                self.export_load_button.config(state='disabled')
                self.export_label.config(text='ERROR: Joined dataframe is null. Please rerun Join Data', fg='red')
        else:
            self.export_fp_button.config(state='normal')
            self.export_label.config(text='Select Export Directory', fg='red')
            self.export_load_button.config(state='disabled')

    def export_path(self):
        fp = filedialog.askdirectory(initialdir='/', title='Select a directory')

        # Sets GUI label text to selected filepath, exits function if not selected
        try:
            if len(fp) != 0:
                os.chdir(fp)
                if len(self.joined_df) != 0:
                    self.export_label.config(text=str(fp), fg='green')
                    self.export_load_button.config(state='normal')
                else:
                    self.export_label.config(text='ERROR: Filepath Selected but no join data to export', fg='red')
                    self.export_load_button.config(state='disabled')
            else:
                self.export_label.config(text='ERROR: No Filepath Selected', fg='red')
                self.export_load_button.config(state='disabled')
                return
        except Exception as fp_error:
            self.export_label.config(text='ERROR: ' + str(fp_error), fg='red')
            self.export_load_button.config(state='disabled')
            return

    def export_data(self):
        fp = self.export_label['text']

        if fp[:6] != 'ERROR:':
            os.chdir(fp)
        else:
            return

        try:

            with pd.ExcelWriter('Site Data Joined.xlsx') as writer:
                export_data = self.joined_df
                export_data.to_excel(writer, sheet_name='Joined Site Data', index=True, freeze_panes=(1, 1))

                workbook = writer.book
                worksheet = writer.sheets['Joined Site Data']

                (max_row, max_col) = export_data.shape

                # Expand index field
                index_format = workbook.add_format({'bold': True})
                worksheet.set_column(0, 0, 20, index_format)

                # Highlight duplicates in index field
                dupe_format = workbook.add_format({'bg_color': '#FF5050'})
                worksheet.conditional_format(0, 0, max_row, 0,
                                             {'type': 'duplicate',
                                              'format': dupe_format})

                # Highlight blank data in red
                blank_format = workbook.add_format({'bg_color': '#FF9999'})
                worksheet.conditional_format(1, 0, max_row, max_col,
                                             {'type': 'blanks',
                                              'format': blank_format})

                # highlight custom generate rows with border
                custom_format = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'right': 2, 'bold': True})
                custom_cols = [col for col in export_data.columns if ' - MERGE' in col
                               or ' AT ' in col or ' M/O ' in col]
                for col in custom_cols:
                    col_position = export_data.columns.get_loc(col) + 1
                    worksheet.set_column(col_position, col_position, None, custom_format)

                # Highlight Primary values by colour scale
                primary_fields = [col for col in export_data.columns if 'Resp. Score' in col or 'Art. Score' in col]
                for col in primary_fields:
                    col_position = export_data.columns.get_loc(col) + 1
                    worksheet.conditional_format(1, col_position, max_row, col_position,
                                                 {'type': '2_color_scale',
                                                  'min_type': 'num',
                                                  'min_value': 1,
                                                  'min_color': '#F4B084',

                                                  'mid_type': 'num',
                                                  'mid_value': 125,
                                                  'mid_color': '#FFE699',

                                                  'max_type': 'num',
                                                  'max_value': 250,
                                                  'max_color': '#A9D08E'
                                                  })

                # Highlight percentage values by colour scale
                tertiary_fields = [col for col in export_data.columns if 'M-L' in col or 'O-L' in col]
                tertiary_values = [col for col in tertiary_fields if 'Test' in col or 'Total' in col]
                for col in tertiary_values:
                    col_position = export_data.columns.get_loc(col) + 1
                    worksheet.conditional_format(1, col_position, max_row, col_position,
                                                 {'type': '2_color_scale',
                                                  'min_type': 'num',
                                                  'min_value': 0,
                                                  'min_color': '#F4B084',

                                                  'mid_type': 'num',
                                                  'mid_value': 100,
                                                  'mid_color': '#FFE699',

                                                  'max_type': 'num',
                                                  'max_value': 200,
                                                  'max_color': '#A9D08E'
                                                  })

                self.export_label.config(text='Data exported as "Site Data Joined.xlsx" to selected folder', fg='green')
                self.export_load_button.config(fg='green')

        except Exception as error:
            self.export_label.config(text='ERROR: ' + str(error), fg='red')
            self.export_load_button.config(state='disabled')
            return
    # endregion


if __name__ == '__main__':
    root = tk.Tk()
    MainFrame(root)
    root.mainloop()
