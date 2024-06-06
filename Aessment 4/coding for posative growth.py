from tkinter import *

# define class App
class App(Tk):
    def __init__(self): # constructor for App
        super().__init__()
        
        self.title("Assessement 4")
        self.geometry("500x300")
        
        ass4_frame = Ass4_Frame(self)
        ass4_frame.show_frame()
    
class Ass4_Frame(Frame):
    def __init__(self, parent): # constructor for frame
        super().__init__(parent)
        
    def show_frame(self): # method to create buttons and textboxes
        self.pack(pady=20)
        
        self.print_all = Button(self, text="Print All", font="Helvetica")
        self.pos_growth = Button(self, text="+ Growth", font="Helvetica")
        self.filt_by_date = Button(self, text="By Date", font="Helvetica")
        self.date_range = Button(self, text="By Date Range", font="Helvetica")
        
        # create text box labels
        self.label_date = Label(self, text="Filter Date").grid(row=4, column = 0, pady=10, padx= 10)
        self.range_date = Label(self, text="Date Range").grid(row=5, column = 0, pady=10, padx= 10)
        
        # create text boxes
        self.text_date = Entry(self, width = 10).grid(row=4, column=1, pady=10, padx= 10)
        self.text_startDate = Entry(self,  width = 10).grid(row=5, column = 1, pady=10, padx= 10)
        self.text_endDate = Entry(self,  width = 10).grid(row=5, column = 2, pady=10, padx= 10)
        
        
        # lay the buttons within the frame
        self.print_all.grid(row=2, column = 0, pady=10, padx= 10)
        self.pos_growth.grid(row=2, column = 1, pady=10, padx= 10)
        self.filt_by_date.grid(row=3, column = 0, pady=10, padx= 10)
        self.date_range.grid(row=3, column=1, pady=10, padx= 10)

# instantiate

app = App()
app.mainloop()
