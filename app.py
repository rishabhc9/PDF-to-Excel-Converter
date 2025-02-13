import subprocess
import tkinter as tk
from tkinter import Button

def load_program1():
    subprocess.Popen(["python", "Grid-Based Table PDF.py"])

def load_program2():
    subprocess.Popen(["python", "Invisible Grid Table PDF.py"])

root = tk.Tk()
root.title("PDF to Excel Conveerter")
root.geometry("300x100")

btn1 = Button(root, text="Convert Grid-Based Table PDF", command=load_program1)
btn1.pack(pady=10)

btn2 = Button(root, text="Convert Invisible Grid Table PDF", command=load_program2)
btn2.pack(pady=10)

root.mainloop()
