#
# simple spreadsheet program
#
import tkinter as tk
import re

class Program:
    def __init__(self, root):
        self.root = root
        self.entry = []
        self.sv = []
        
        root.title("Untitled: spreadsheet")
        root.geometry("650x490")
        
        # add a menu: ..................................................
        menubar = tk.Menu(root)
        filemenu = tk.Menu(menubar,tearoff=0)
        filemenu.add_command(label="New", command=self.new_file)
        filemenu.add_command(label="Open", command=self.open_file)
        filemenu.add_command(label="Save", command=self.save_file)
        filemenu.add_command(label="Save as", command=self.save_as_file)
        filemenu.add_command(label="Exit", command=self.exit_menu)
        menubar.add_cascade(label="File", menu=filemenu)
        #---------------------------------------------------------------
        editmenu = tk.Menu(menubar,tearoff=0)
        editmenu.add_command(label="Cut", command=self.cut)
        editmenu.add_command(label="Copy", command=self.copy)
        editmenu.add_command(label="Paste", command=self.paste)
        menubar.add_cascade(label="Edit", menu=editmenu)
        #---------------------------------------------------------------
        optionmenu = tk.Menu(menubar,tearoff=0)
        optionmenu.add_command(label="Font", command=self.changefont)
        optionmenu.add_command(label="Font Size", command=self.changefontsize)
        optionmenu.add_command(label="Font Style", command=self.changefontstyle)
        menubar.add_cascade(label="Options", menu=optionmenu)
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
        self.frame.bind("<Configure>", self.onFrameConfigure)
        self.create_grid(3, 5)
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
        self.child_window_new.geometry("269x116")                #setting window size
        self.child_window_new.resizable(width=True, height=True)

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
        # get new grid data and check that it is correct
        
        # remove current grid
        self.delete_grid()
        #print(len(self.entry[0]))
        #print(len(self.sv[0]))
        #self.entry_x.destroy()

        # add new grid
        
        # close this window
        #self.cancel_new_spreadsheet()
        print("command")
    def cancel_new_spreadsheet(self):
        self.child_window_new.destroy()
        self.child_window_new.update()
    #--------------------------------------------------------------------
    def cut(self):
        pass
    def copy(self):
        pass
    def paste(self):
        pass
    #--------------------------------------------------------------------
    def open_file(self):
        pass
    def save_file(self):
        pass
    def save_as_file(self):
        pass
    def exit_menu(self):
        pass
    #--------------------------------------------------------------------
    def delete_grid(self):
        # 1st delete all string var
        for n in reversed(self.sv):
            for nn in reversed(n):
                nn.set('')
        # 2nd delete all entry
        for n in reversed(self.entry):
            for nn in reversed(n):
                nn.destroy()
    def create_grid(self, width, height):
        for i in range(width):
            self.entry.append([])
            self.sv.append([])
            for c in range(height):
                self.sv[i].append(tk.StringVar())
                self.sv[i][c].trace("w", lambda name, index, mode, sv=self.sv[i][c], i=i, c=c: self.callback(sv, i, c))
                self.entry[i].append(tk.Entry(self.frame, textvariable=self.sv[i][c]).grid(row=c, column=i))
    def onFrameConfigure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    def callback(self, sv, column, row):
        print("Column: " + str(column) + ", Row: " + str(row) + " = " + sv.get())
    #--------------------------------------------------------------------
    def changefont(self):
        pass
    def changefontsize(self):
        pass
    def changefontstyle(self):
        pass

def main():
    root = tk.Tk()
    Program(root)
    root.mainloop()
 
if __name__ == '__main__':
    main()
