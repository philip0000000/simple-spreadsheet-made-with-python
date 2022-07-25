#
# simple spreadsheet program
# Written by philip0000000
# Find the project here [https://github.com/philip0000000/simple-spreadsheet-made-with-python]
#
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import re

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
        
        # add a menu: ..................................................
        menubar = tk.Menu(root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="New", command=self.new_file)
        filemenu.add_command(label="Open", command=self.open_file)
        filemenu.add_command(label="Save", command=self.save_file)
        filemenu.add_command(label="Save as", command=self.save_as_file)
        filemenu.add_command(label="Exit", command=self.exit_program)
        menubar.add_cascade(label="File", menu=filemenu)
        #---------------------------------------------------------------
        editmenu = tk.Menu(menubar, tearoff=0)
        editmenu.add_command(label="Cut", command=self.cut)
        editmenu.add_command(label="Copy", command=self.copy)
        editmenu.add_command(label="Paste", command=self.paste)
        menubar.add_cascade(label="Edit", menu=editmenu)
        #---------------------------------------------------------------
        root.config(menu=menubar)
        
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
    #--------------------------------------------------------------------
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
    def open_file(self):
        filename = str(tk.filedialog.askopenfilename(title="Open File", filetypes=[("File", "*"), ("Text Document", ".txt")]))
        if len(filename) > 0:
            try:
                # delete current spreadsheet
                self.delete_spreadsheet()
                
                file_data = []
                with open(filename) as f:
                    for line in f:
                        # Extract substrings between square brackets, using regex
                        res = re.findall(r'\[.*?\]', line)
                        file_data.append(res)
                
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
                            if len(file_data[i][j]) > 2:
                                self.sv[j][i].set((file_data[i][j])[1:-1])
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
                    cell_data = "[" + self.sv[j][i].get() + "]"
                    # add space, except for 1st cell
                    if (j != 0):
                        cell_data = " " + cell_data
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
    def save_as_file(self):
        filename = str(tk.filedialog.asksaveasfilename(title="Save as File", defaultextension=".txt", filetypes=[("Text Document", ".txt"), ("File", "*")]))
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
                    cell_data = "[" + self.sv[j][i].get() + "]"
                    # add space, except for 1st cell
                    if (j != 0):
                        cell_data = " " + cell_data
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
            self.entry.append([])
            self.sv.append([])
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

class not_enough_data_error(Exception):
    """Could not find enougth data"""
    pass

def main():
    root = tk.Tk()
    spreadsheet_program(root)
    root.mainloop()

if __name__ == '__main__':
    main()
