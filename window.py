# Import Required Library
from tkinter import *
from tkcalendar import Calendar
from patches_new import build_excel

# Create Object
root = Tk()

# Set geometry
root.geometry("400x400")

# Add Calendar
cal = Calendar(root, selectmode = 'day',date_pattern='y-mm-dd')

cal.pack(pady = 20)

def grad_date():
	date.config(text = "Selected Date is: " + cal.get_date())
	print(cal.get_date())
	print(varScreenshot.get())
	build_excel(cal.get_date(),varScreenshot.get())

# Add toggle screenshots
varScreenshot = BooleanVar()
Checkbutton(root,text="Take screenshots",variable=varScreenshot).pack(pady=20)

# Add Button and Label
Button(root, text = "Get Patches",
	command = grad_date).pack(pady = 20)

date = Label(root, text = "")
date.pack(pady = 20)

# Execute Tkinter
root.mainloop()
