import os
import tkinter as tk
from tkinter import ttk
from openpyxl import Workbook, load_workbook


app = tk.Tk()
app.title("Score Tracker")
app.geometry()
app.resizable(False, False)

frame = tk.Frame(app)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Score Tracker")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

label_frame = tk.Frame(widgets_frame)
label_frame.grid(row=0, column=0)

entry_name = ttk.Entry(label_frame)
entry_name.insert(0, "Name")
entry_name.bind("<FocusIn>", lambda e: entry_name.delete("0", "end") if entry_name.get() == "Name" else None)
entry_name.bind("<FocusOut>", lambda e: entry_name.insert(0, "Name") if entry_name.get() == "" else None)
entry_name.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

entry_course = ttk.Entry(label_frame)
entry_course.insert(0, "Course")
entry_course.bind("<FocusIn>", lambda e: entry_course.delete("0", "end") if entry_course.get() == "Course" else None)
entry_course.bind("<FocusOut>", lambda e: entry_course.insert(0, "Course") if entry_course.get() == "" else None)
entry_course.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

entry_score = ttk.Entry(label_frame)
entry_score.insert(0, "Score")
entry_score.bind("<FocusIn>", lambda e: entry_score.delete("0", "end") if entry_score.get() == "Score" else None)
entry_score.bind("<FocusOut>", lambda e: entry_score.insert(0, "Score") if entry_score.get() == "" else None)
entry_score.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")

button = ttk.Button(label_frame, text="Submit", )
button.grid(row=3, column=0, sticky="nsew")


app.mainloop()
