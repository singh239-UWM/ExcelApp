#libs
from tkinter import filedialog
  
# Function for opening the
# file explorer window
def browseFiles():
    filepath = filedialog.askopenfilename(initialdir = "/", 
                                          title = "Select a File", 
                                          filetypes = (("Microsoft Excel Worksheet", "*.xlsx*"), ("All Files", "*.*")) )
    # Change label contents
    
    return filepath