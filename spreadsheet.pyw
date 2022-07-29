#
# simple spreadsheet program
# Written by philip0000000
# Find the project here [https://github.com/philip0000000/simple-spreadsheet-made-with-python]
#
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class spreadsheet_program:
    """simple spreadsheet program"""
    filename = "Untitled"
    file_changed = False
    # cells and data of spreadsheet
    entry = []
    sv = []
    def __init__(self, root):
        self.root = root
        self.root.title("spreadsheet: " + self.filename)
        root.geometry("650x490")
        
        # add a menu bar
        menu_bar = tk.Menu(root)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="New", command=self.new_file)
        file_menu.add_command(label="Open", command=self.open_file)
        file_menu.add_command(label="Save", command=self.save_file)
        file_menu.add_command(label="Save as", command=self.save_as_file)
        file_menu.add_command(label="Exit", command=self.exit_program)
        menu_bar.add_cascade(label="File", menu=file_menu)
        
        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="Cut", command=self.cut)
        edit_menu.add_command(label="Copy", command=self.copy)
        edit_menu.add_command(label="Paste", command=self.paste)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)
        
        table_menu = tk.Menu(menu_bar, tearoff=0)
        table_menu.add_command(label="Add vertical table", command=self.add_vertical_table)
        table_menu.add_command(label="Remove vertical table", command=self.remove_vertical_table)
        table_menu.add_command(label="Add horizontal table", command=self.add_horizontal_table)
        table_menu.add_command(label="Remove horizontal table", command=self.remove_horizontal_table)
        menu_bar.add_cascade(label="Table", menu=table_menu)
        
        root.config(menu=menu_bar)
        
        self.canvas = tk.Canvas(self.root, background="#ffffff", borderwidth=0)
        self.frame = tk.Frame(self.canvas, background="#ffffff")
        self.scrolly = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollx = tk.Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.scrolly.set)#, xscrollcommand=self.scrollx.set)
        self.canvas.create_window((4,4), window=self.frame, anchor="nw", tags="self.frame")
        self.scrolly.pack(side="left", fill="y")
        self.canvas.pack(side="top", fill="both", expand=True)
        self.scrollx.pack(side="bottom", fill="x")
        self.frame.bind("<Configure>", self.on_frame_configure)
        self.create_spreadsheet(3, 5)
    
    def limit_string_var_to_four(self, var):
        # remove everything except for numbers
        value = var.get()
        value = re.sub('\D', '', value)
        var.set(value)
        # max 4 numbers
        value = var.get()
        if len(value) > 4: var.set(value[:4])
    def validate_on_float(self, action, index, value_if_allowed,
                          prior_value, text, validation_type, trigger_type, widget_name):
        if value_if_allowed:
            try:
                float(value_if_allowed)
                return True
            except ValueError:
                return False
        else:
            return False
    def new_file(self):
        # create child window when new spreadsheet is required
        self.child_window_new = tk.Toplevel(self.root)
        self.child_window_new.title("spreadsheet")               # setting title
        self.child_window_new.geometry("269x116")                # setting window size
        self.child_window_new.resizable(width=True, height=True)
        self.child_window_new.transient(self.root)               # top of the main window

        # Add labels
        input_x_and_y_text=tk.Label(self.child_window_new, text="New x and y dimension for spreadsheet")
        input_x_and_y_text.place(x=0, y=0, width=236, height=30)
        x_text=tk.Label(self.child_window_new, text="X")
        x_text.place(x=0, y=30, width=30, height=49)
        y_text=tk.Label(self.child_window_new, text="Y")
        y_text.place(x=0, y=60, width=30, height=49)

        # Add entry
        #vcmd = (self.root.register(self.validate_on_float),
                 #'%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        #entry_x=tk.Entry(self.child_window_new, text="entry_x", validate = 'key', validatecommand = vcmd)
        #entry_x.place(x=30, y=40, width=35, height=30)
        self.sv_x = tk.StringVar()
        self.sv_x.trace('w', lambda *_, var=self.sv_x: self.limit_string_var_to_four(var))
        self.entry_x=tk.Entry(self.child_window_new, text="entry_x", textvariable=self.sv_x)
        self.entry_x.place(x=30, y=40, width=35, height=30)
        self.sv_y = tk.StringVar()
        self.sv_y.trace('w', lambda *_, var=self.sv_y: self.limit_string_var_to_four(var))
        entry_y=tk.Entry(self.child_window_new, text="entry_y", textvariable=self.sv_y)
        entry_y.place(x=30, y=70, width=35, height=30)

        # Add buttons
        new_button=tk.Button(self.child_window_new, text="New", command=self.new_spreadsheet)
        new_button.place(x=120, y=70, width=68, height=30)
        cancel_button=tk.Button(self.child_window_new, text="Cancel", command=self.cancel_new_spreadsheet)
        cancel_button.place(x=190, y=70, width=68, height=30)
    def new_spreadsheet(self):
        # get new spreadsheet data and check that it is correct
        x_val = self.sv_x.get()
        y_val = self.sv_y.get()
        try:
            x = int(x_val)
            y = int(y_val)
            
            # delete old spreadsheet
            self.delete_spreadsheet()
            
            # add new spreadsheet
            self.create_spreadsheet(x, y)
            
            # close this window
            self.cancel_new_spreadsheet()
            
            # set title
            self.filename = "Untitled"
            self.file_changed = False
            self.root.title("spreadsheet: " + self.filename)
        except Exception as e:
            messagebox.showwarning("error", "X and/or Y data was wrong")
    def cancel_new_spreadsheet(self):
        self.child_window_new.destroy()
        self.child_window_new.update()
    # https://en.wikipedia.org/wiki/Comma-separated_values#Basic_rules
    def open_file(self):
        filename = str(tk.filedialog.askopenfilename(title="Open File", filetypes=[("File", "*"), ("Text Document", ".txt"), ("Comma-separated values", ".csv")]))
        if len(filename) > 0:
            try:
                # delete current spreadsheet
                self.delete_spreadsheet()
                
                # get file data
                file_data = []
                with open(filename) as f:
                    for raw_line_data in f:
                        line_data = []
                        # if two double quotation marks follow eachother, make theme only one
                        processed_line_data = ""
                        n = 0
                        raw_line_data_length = len(raw_line_data)
                        while n < raw_line_data_length:
                            if n+3 < raw_line_data_length:
                                if raw_line_data[n] == "," and raw_line_data[n+1] == "\"" and raw_line_data[n+2] == "\"" and raw_line_data[n+3] == ",":
                                    processed_line_data += ","
                                    n += 3
                            # turn all 2 double quotation mark into 1 double quotation mark
                            if raw_line_data[n] == "\"" and raw_line_data[n+1] == "\"":
                                # if double quotation mark looks like this: ",\"\","
                                # do nothing
                                if n > 0:
                                    if raw_line_data[n-1] != ",":
                                        processed_line_data += "\""
                                        n += 2
                                    else:
                                        processed_line_data += raw_line_data[n]
                                        n += 1
                                else:
                                    processed_line_data
                                    n += 1
                            else:
                                processed_line_data += raw_line_data[n]
                                n += 1
                        # Extract substrings between comma
                        cell_data = ""
                        inside_double_quotation_mark = False
                        n = 0
                        processed_line_data_length = len(processed_line_data)
                        while n < processed_line_data_length:
                            # if we get ",\"" we set flag and jump 2 step forward
                            if n+1 < processed_line_data_length and processed_line_data[n] == "," and processed_line_data[n+1] == "\"" and inside_double_quotation_mark == False:
                                line_data.append(cell_data)
                                cell_data = ""
                                inside_double_quotation_mark = True
                                n += 2
                            # if we get "\"," we make flag false and jump 3 steps forward
                            elif n+1 < processed_line_data_length and processed_line_data[n] == "\"" and processed_line_data[n+1] == "," and inside_double_quotation_mark == True:
                                inside_double_quotation_mark = False
                                line_data.append(cell_data)
                                cell_data = ""
                                n += 2
                            # if we get comma with no flag, we save cell data and add new cell data
                            elif processed_line_data[n] == "," and inside_double_quotation_mark == False:
                                # new cell, save current cell data and start obtaining the new data from the new cell
                                line_data.append(cell_data)
                                cell_data = ""
                                n += 1
                            else:
                                # do not add new line escape character
                                if processed_line_data[n] != "\n":
                                    cell_data += processed_line_data[n]
                                n += 1
                        # if something is left in the cell data, add it
                        if len(cell_data) > 0:
                            line_data.append(cell_data)
                        # add line data to file data
                        file_data.append(line_data)
                
                # get height and width value
                height = len(file_data)
                if height < 1:
                    raise not_enough_data_error
                width = 0
                # loop throught every list to get biggest value
                for n in file_data:
                    if len(n) > width:
                        width = len(n)
                if width < 1:
                    raise not_enough_data_error
                    
                # create spreadsheet
                self.create_spreadsheet(width, height)
                
                # add value from file to spreadsheet
                i = 0
                while i < height:
                    j = 0
                    while j < width:
                        # check if 2D index exist before adding to spreadsheet
                        jj = len(file_data[i])
                        if j < jj:
                            self.sv[j][i].set(file_data[i][j])
                        j += 1
                    i += 1
                
                # set title
                self.filename = filename
                self.file_changed = False
                self.root.title("spreadsheet: " + self.filename)
            except IOError:
                tk.tkMessageBox.showwarning("Open file","Cannot open this file...")
            except not_enough_data_error:
                messagebox.showwarning("error", "not enough data in file")
                # set title
                self.filename = "Untitled"
                file_changed = False
                self.root.title("spreadsheet: " + self.filename)
    def delete_spreadsheet(self):
        # 1st delete all string var
        for n in reversed(self.sv):
            for nn in reversed(n):
                nn.set('')
            n.clear()
        self.sv.clear()
        # 2nd delete all entry
        for n in reversed(self.entry):
            for nn in reversed(n):
                nn.destroy()
            n.clear()
        self.entry.clear()
    def create_spreadsheet(self, width, height):
        for i in range(width):
            self.sv.append([])
            self.entry.append([])
            for c in range(height):
                self.sv[i].append(tk.StringVar())
                self.sv[i][c].trace("w", lambda name, index, mode, sv=self.sv[i][c], i=i, c=c: self.callback(i, c))
                self.entry[i].append(tk.Entry(self.frame, textvariable=self.sv[i][c]))
                self.entry[i][c].grid(row=c, column=i)
    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    def callback(self, column, row):
        # set title
        file_changed = True
        self.root.title("spreadsheet: " + self.filename + "*")
    # https://en.wikipedia.org/wiki/Comma-separated_values#Basic_rules
    def save_file(self):
        if self.filename == "Untitled":
            self.save_as_file()
        else:
            column_length = len(self.sv[0])
            row_length = len(self.sv)
            
            # open file
            f = open(self.filename, "wb")
            
            # write to file
            i = 0
            while i < column_length:
                j = 0
                while j < row_length:
                    raw_cell_data = self.sv[j][i].get()
                    cell_data = ""
                    # if the data in the cell has double quotation marks add a double quotation marks before it
                    for c in raw_cell_data:
                        if c == "\"":
                            cell_data += "\"\""
                        else:
                            cell_data += c
                    # if the data in the cell has comma, surround the data with double quotation marks
                    if "," in raw_cell_data:
                        cell_data = "\"" + cell_data + "\""
                    # add comma, except for 1st cell
                    if (j != 0):
                        cell_data = "," + cell_data
                    cell_data = cell_data.encode("utf-8")
                    f.write(cell_data)
                    j += 1
                i += 1
                # add new line, except for last line
                if i < column_length:
                    f.write("\n".encode("utf-8"))
            
            # close file
            f.close()
            
            # set title
            file_changed = False
            self.root.title("spreadsheet: " + self.filename)
    # https://en.wikipedia.org/wiki/Comma-separated_values#Basic_rules
    def save_as_file(self):
        filename = str(tk.filedialog.asksaveasfilename(title="Save as File", defaultextension=".txt", filetypes=[("Text Document", ".txt"), ("Comma-separated values", ".csv"), ("File", "*")]))
        if len(filename) > 0:
            column_length = len(self.sv[0])
            row_length = len(self.sv)
            
            # open file
            f = open(filename, "wb")
            
            # write to file
            i = 0
            while i < column_length:
                j = 0
                while j < row_length:
                    raw_cell_data = self.sv[j][i].get()
                    cell_data = ""
                    # if the data in the cell has double quotation marks add a double quotation marks before it
                    for c in raw_cell_data:
                        if c == "\"":
                            cell_data += "\"\""
                        else:
                            cell_data += c
                    # if the data in the cell has comma, surround the data with double quotation marks
                    if "," in raw_cell_data:
                        cell_data = "\"" + cell_data + "\""
                    # add comma, except for 1st cell
                    if (j != 0):
                        cell_data = "," + cell_data
                    cell_data = cell_data.encode("utf-8")
                    f.write(cell_data)
                    j += 1
                i += 1
                # add new line, except for last line
                if i < column_length:
                    f.write("\n".encode("utf-8"))
            
            # close file
            f.close()
            
            # set title
            self.filename = filename
            self.file_changed = False
            self.root.title("spreadsheet: " + self.filename)
    def exit_program(self):
        if self.file_changed:
            if tk.tkMessageBox.askyesno("Quit", "Do you want to save the file?"):
                if self.filename == "Untitled":
                    self.save_as_file()
                else:
                    self.save_file()
        self.root.destroy()
    #--------------------------------------------------------------------
    def cut(self):
        try:
            self.copy()
            widget = self.root.focus_get()
            widget.delete("sel.first", "sel.last")
            # set title
            self.file_changed = True
            self.root.title("spreadsheet: " + self.filename + "*")
        except tk.TclError:
            pass
    def copy(self):
        try:
            widget = self.root.focus_get()
            widget.clipboard_clear()
            text = widget.selection_get()
            widget.clipboard_append(text)
        except tk.TclError:
            pass
    def paste(self):
        try:
            widget = self.root.focus_get()
            text = widget.selection_get(selection="CLIPBOARD")
            widget.insert(tk.INSERT, text)
            # set title
            self.file_changed = True
            self.root.title("spreadsheet: " + self.filename + "*")
        except tk.TclError:
            pass
    #--------------------------------------------------------------------
    def add_vertical_table(self):
        if len(self.sv) > 0:
            # get number of vertical cells
            v_len = len(self.sv[0])
            h_len = len(self.sv)
            
            # add new vertical list to the 2D list
            self.sv.append([])
            self.entry.append([])
            
            # add cells vertical
            for n in range(v_len):
                self.sv[-1].append(tk.StringVar())
                self.sv[-1][n].trace("w", lambda name, index, mode, sv=self.sv[-1][n], i=h_len, c=n: self.callback(i, c))
                self.entry[-1].append(tk.Entry(self.frame, textvariable=self.sv[-1][n]))
                self.entry[-1][n].grid(row=n, column=h_len)
        else:
            # add only 1 cell
            # add new vertical list to the 2D list
            self.sv.append([])
            self.entry.append([])
            
            self.sv[0].append(tk.StringVar())
            self.sv[0][0].trace("w", lambda name, index, mode, sv=self.sv[0][0], i=0, c=0: self.callback(i, c))
            self.entry[0].append(tk.Entry(self.frame, textvariable=self.sv[0][0]))
            self.entry[0][0].grid(row=0, column=0)
    def remove_vertical_table(self):
        if len(self.sv) > 0:
            # get number of vertical cells
            v_len = len(self.sv[0])
            h_len = len(self.sv)
            h_len -= 1
            
            # check if there exist data, if yes, as for confirmation
            exist_data = False
            for n in range(v_len):
                if self.sv[-1][n].get() != '':
                    # data has been found
                    exist_data = True
                    break
            
            msg_box_answer = "yes"
            if exist_data == True:
                msg_box_answer = tk.messagebox.askquestion("Delete data?", "Are you sure you want to delete the vertical table with data?", icon = "warning")
            
            if msg_box_answer == "yes":
                # 1st delete all string var
                for n in range(v_len):
                    self.sv[-1][n].set('')
                self.sv[-1].clear()
                del self.sv[-1]
                # 2nd delete all entry
                for n in range(v_len):
                    self.entry[-1][n].destroy()
                self.entry[-1].clear()
                del self.entry[-1]
    def add_horizontal_table(self):
        if len(self.sv) > 0:
            # get number of horizontal cells
            h_len = len(self.sv)
            v_len = len(self.sv[0])
            
            # add cells vertical
            for n in range(h_len):
                self.sv[n].append(tk.StringVar())
                self.sv[n][v_len].trace("w", lambda name, index, mode, sv=self.sv[n][v_len], i=n, c=v_len: self.callback(i, c))
                self.entry[n].append(tk.Entry(self.frame, textvariable=self.sv[n][v_len]))
                self.entry[n][v_len].grid(row=v_len, column=n)
        else:
            # add only 1 cell
            # add new vertical list to the 2D list
            self.sv.append([])
            self.entry.append([])
            
            self.sv[0].append(tk.StringVar())
            self.sv[0][0].trace("w", lambda name, index, mode, sv=self.sv[0][0], i=0, c=0: self.callback(i, c))
            self.entry[0].append(tk.Entry(self.frame, textvariable=self.sv[0][0]))
            self.entry[0][0].grid(row=0, column=0)
    def remove_horizontal_table(self):
        if len(self.sv) > 0:
            # get number of vertical cells
            h_len = len(self.sv)
            
            # check if there exist data, if yes, as for confirmation
            exist_data = False
            for n in range(h_len):
                if  self.sv[n][-1].get() != '':
                    # data has been found
                    exist_data = True
                    break
            
            msg_box_answer = "yes"
            if exist_data == True:
                msg_box_answer = tk.messagebox.askquestion("Delete data?", "Are you sure you want to delete the horizontal table with data?", icon = "warning")
            
            if msg_box_answer == "yes":
                # 1st delete all string var
                for n in range(h_len):
                    self.sv[n][-1].set('')
                    del self.sv[n][-1]
                # 2nd delete all entry
                for n in range(h_len):
                    self.entry[n][-1].destroy()
                    del self.entry[n][-1]
                # delete everything if everything is empty
                if len(self.sv[0]) == 0:
                    self.sv.clear()
                    self.entry.clear()

class not_enough_data_error(Exception):
    """Could not find enougth data"""
    pass

def main():
    root = tk.Tk()
    spreadsheet_program(root)
    root.mainloop()

if __name__ == '__main__':
    main()
