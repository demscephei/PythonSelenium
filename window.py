# Import Required Libraries
import tkinter as tk
from tkinter import ttk
import tkcalendar as tkc
from tkcalendar import Calendar
from patches import build_excel
import threading
import queue

def update_ui():
	try:
		while True:
			progress, status = progressQueue.get_nowait()
			cal['state'] = "disabled"
			btCreate['state'] = "disabled"
			cbScreenshot['state'] = "disabled"
			progress_bar['value'] = progress
			lbStatus.config(text=status)
			root.update_idletasks()
			if progress >= 100:
				lbStatus.config(text="Patches Spreadsheet created.")
				cal['state'] = "normal"
				btCreate['state'] = "normal"
				cbScreenshot['state'] = "normal"
				root.bell()
				break
	except queue.Empty:
		root.after(100, update_ui)

def start_task():
	taskThread = threading.Thread(target=build_excel, args=(cal.get_date(),varScreenshot.get(),progressQueue,))
	taskThread.start()
	root.after(100,update_ui)

# Create Window
root = tk.Tk()

# Set geometry
root.geometry("400x460")

root.title("Patch Spreadsheet Creator")
lbTitle = tk.Label(root,text="Patch Spreadsheet Creator").pack(pady=10)
lbInstructions = tk.Label(root,text="1. Select previous patch wednesday date.\n2.Toggle screenshots (optional).\n3. Create Patches Spreadsheet!").pack(pady=2)

# Add Calendar
cal = tkc.Calendar(root, selectmode = 'day',date_pattern='y-mm-dd')
cal.pack(pady = 10)

# Add toggle screenshots
varScreenshot = tk.BooleanVar()
cbScreenshot = tk.Checkbutton(root,text="Take screenshots",variable=varScreenshot)
cbScreenshot.pack(pady=5)

# Add progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal",length=300,mode="determinate",phase=10)
progress_bar.pack(pady=2)
lbStatus = tk.Label(root,text="Click Get Patches to Begin")
lbStatus.pack(pady=5)

# Add Button and Label
btCreate = tk.Button(root, text = "Create Spreadsheet",command = start_task)
btCreate.pack(pady = 20)
progressQueue = queue.Queue()

# Execute Tkinter
root.mainloop()
